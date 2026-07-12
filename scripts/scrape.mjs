// Govt jobs scraper: fetch official portals, detect changes, extract jobs via Gemini,
// diff against data/jobs.json. Designed for a 15-minute GitHub Actions cron.
//
// Usage:
//   node scripts/scrape.mjs                 run all sources that are due
//   node scripts/scrape.mjs --force         ignore cadence + content-hash gating
//   node scripts/scrape.mjs --only drdo-rac run a single source
//
// Gemini free-tier budget: calls happen only when a source's page content hash
// changes (HTML sources) or its grounded cadence elapses, capped per run.

import { createHash } from "node:crypto";
import { readFileSync, writeFileSync, existsSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = join(dirname(fileURLToPath(import.meta.url)), "..");
const DATA_DIR = join(ROOT, "data");
const JOBS_FILE = join(DATA_DIR, "jobs.json");
const STATE_FILE = join(DATA_DIR, "state.json");

const MODEL = process.env.GEMINI_MODEL || "gemini-2.5-flash";
const API_BASE = "https://generativelanguage.googleapis.com/v1beta/models";
const MAX_GEMINI_CALLS_PER_RUN = 6;
const FETCH_TIMEOUT_MS = 45_000;
const MAX_PAGE_CHARS = 18_000;
const NEW_BADGE_HOURS = 48;

const PROFILE = `Candidate profile: B.Tech in Mechanical Engineering, currently working as a
software engineer in the private sector with 2-3 years of experience. Interested primarily in
MANAGEMENT and IT roles in Indian government / PSU / banking / regulators / state PSC — not core
mechanical engineering. IMPORTANT eligibility rule: most government IT posts (IBPS SO IT, SBI SCO
tech, NIC Scientist-B) require a CS/IT/Electronics DEGREE, which this candidate does not have, so
mark those "gated". Any-graduate management exams (PCS, RBI Grade B, NABARD Grade A, bank PO,
SEBI Grade A General which accepts any engineering bachelor's) are "strong". Core engineering
posts where Mechanical is eligible are "moderate". Irrelevant posts (medical, law, teaching,
constable etc.) are "low".`;

const JOB_SCHEMA = {
  type: "ARRAY",
  items: {
    type: "OBJECT",
    properties: {
      org: { type: "STRING", description: "Recruiting organisation, short form, e.g. 'UPSC · Central Ministries'" },
      title: { type: "STRING", description: "Post / exam name" },
      summary: { type: "STRING", description: "Max 45 words: what the post is, vacancies, eligibility, selection stages" },
      advtNo: { type: "STRING", nullable: true },
      posts: { type: "STRING", nullable: true, description: "Vacancy count as short text, e.g. '745 posts'" },
      url: { type: "STRING", description: "Absolute link to the official notice or portal" },
      status: { type: "STRING", enum: ["open", "upcoming", "closed"] },
      opensOn: { type: "STRING", nullable: true, description: "ISO date YYYY-MM-DD or null" },
      closesOn: { type: "STRING", nullable: true, description: "ISO date YYYY-MM-DD or null" },
      expected: { type: "STRING", nullable: true, description: "For upcoming posts: expected window as text, e.g. 'Jul-Aug 2026'" },
      fit: { type: "STRING", enum: ["strong", "moderate", "gated", "low"] },
      fitLabel: { type: "STRING", description: "Short chip text, e.g. 'Strong fit — management' or 'Degree-gated: CS/IT'" },
      tags: { type: "ARRAY", items: { type: "STRING" }, description: "0-2 short facts like 'Any graduate'" }
    },
    required: ["org", "title", "summary", "url", "status", "fit", "fitLabel"]
  }
};

function loadEnvFile() {
  const envPath = join(ROOT, ".env");
  if (!existsSync(envPath)) return;
  for (const line of readFileSync(envPath, "utf8").split(/\r?\n/)) {
    const m = line.match(/^\s*([A-Z_][A-Z0-9_]*)\s*=\s*(.*)\s*$/);
    if (m && !process.env[m[1]]) process.env[m[1]] = m[2];
  }
}

function loadJSON(file, fallback) {
  try { return JSON.parse(readFileSync(file, "utf8")); } catch { return fallback; }
}

function sha256(text) {
  return createHash("sha256").update(text).digest("hex");
}

function slug(text) {
  return text.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "").slice(0, 80);
}

function todayISO() {
  return new Date().toISOString().slice(0, 10);
}

function htmlToText(html) {
  return html
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<!--[\s\S]*?-->/g, " ")
    .replace(/<(br|\/p|\/div|\/tr|\/li|\/h[1-6])[^>]*>/gi, "\n")
    .replace(/<a\s[^>]*href="([^"]+)"[^>]*>([\s\S]*?)<\/a>/gi, "$2 [link: $1]")
    .replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/g, " ").replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">")
    .replace(/&#\d+;|&[a-z]+;/gi, " ")
    .replace(/[ \t]+/g, " ")
    .replace(/\n\s*\n[\s\n]*/g, "\n")
    .trim()
    .slice(0, MAX_PAGE_CHARS);
}

async function fetchPage(url) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), FETCH_TIMEOUT_MS);
  try {
    const res = await fetch(url, {
      signal: controller.signal,
      redirect: "follow",
      headers: {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml,application/json;q=0.9,*/*;q=0.8"
      }
    });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    return await res.text();
  } finally {
    clearTimeout(timer);
  }
}

async function geminiRequest(body, attempt = 1) {
  const res = await fetch(`${API_BASE}/${MODEL}:generateContent?key=${process.env.GEMINI_API_KEY}`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body)
  });
  if (res.status === 429 || res.status === 503) {
    if (attempt >= 3) throw new Error(`Gemini ${res.status} after ${attempt} attempts`);
    const wait = 15_000 * attempt;
    console.log(`  Gemini ${res.status}, retrying in ${wait / 1000}s...`);
    await new Promise(r => setTimeout(r, wait));
    return geminiRequest(body, attempt + 1);
  }
  if (!res.ok) throw new Error(`Gemini HTTP ${res.status}: ${(await res.text()).slice(0, 300)}`);
  const data = await res.json();
  const text = data.candidates?.[0]?.content?.parts?.map(p => p.text || "").join("") || "";
  if (!text) throw new Error("Gemini returned no text");
  return text;
}

const EXTRACT_INSTRUCTIONS = `You extract Indian government job postings for a job tracker.
${PROFILE}
Today's date: ${todayISO()}.
Rules:
- Only include actual recruitment notices (open, announced/upcoming, or closed within the last 21 days). Ignore results, admit cards, answer keys, interview schedules, tenders.
- Dates in the source may be DD-MM-YYYY or DD.MM.YYYY — convert to ISO YYYY-MM-DD.
- status: "open" if the application window includes today, "upcoming" if it starts later or is only announced, "closed" if it ended.
- Make urls absolute using the source URL's origin when relative.
- If the page shows nothing job-related, return [].`;

async function extractFromContent(source, pageText) {
  const text = await geminiRequest({
    systemInstruction: { parts: [{ text: EXTRACT_INSTRUCTIONS }] },
    contents: [{
      role: "user",
      parts: [{ text: `Source: ${source.name} (${source.url})\n\nPage content:\n${pageText}` }]
    }],
    generationConfig: {
      responseMimeType: "application/json",
      responseSchema: JOB_SCHEMA,
      temperature: 0.1
    }
  });
  return JSON.parse(text);
}

async function extractGrounded(source) {
  // Google Search grounding cannot be combined with responseSchema, so we ask
  // for JSON in the prompt and parse the first array in the reply.
  const text = await geminiRequest({
    contents: [{
      role: "user",
      parts: [{
        text: `${EXTRACT_INSTRUCTIONS}\n\nUse Google Search to answer this:\n${source.query}\n\nReply with ONLY a JSON array matching this shape (no markdown fences):\n[{"org":"","title":"","summary":"","advtNo":null,"posts":null,"url":"","status":"open|upcoming|closed","opensOn":null,"closesOn":null,"expected":null,"fit":"strong|moderate|gated|low","fitLabel":"","tags":[]}]`
      }]
    }],
    tools: [{ google_search: {} }],
    generationConfig: { temperature: 0.1 }
  });
  const start = text.indexOf("[");
  const end = text.lastIndexOf("]");
  if (start === -1 || end <= start) return [];
  return JSON.parse(text.slice(start, end + 1));
}

const STOP_WORDS = new Set(["the", "of", "and", "for", "a", "an", "exam", "examination", "recruitment", "post", "posts", "vacancy", "vacancies", "notification"]);

function tokens(text) {
  return new Set(
    String(text || "").toLowerCase().replace(/[^a-z0-9 ]+/g, " ").split(/\s+/)
      .filter(w => w && !STOP_WORDS.has(w))
  );
}

// Overlap coefficient: intersection / size of the smaller set.
function overlap(a, b) {
  if (!a.size || !b.size) return 0;
  let inter = 0;
  for (const t of a) if (b.has(t)) inter++;
  return inter / Math.min(a.size, b.size);
}

// Stream/discipline words that distinguish otherwise same-named postings, e.g.
// "HAL Management Trainee (Mechanical)" vs "HAL Management Trainee (Computer Science)".
const DISCIPLINES = new Set(["mechanical", "electrical", "civil", "computer", "science", "electronics",
  "it", "chemical", "metallurgical", "mining", "instrumentation", "finance", "hr", "human",
  "law", "legal", "marketing", "agriculture", "rajbhasha", "medical", "telecom"]);

function disciplineConflict(ta, tb) {
  const da = [...ta].filter(t => DISCIPLINES.has(t));
  const db = [...tb].filter(t => DISCIPLINES.has(t));
  return da.length > 0 && db.length > 0 && !da.some(t => tb.has(t));
}

// Same posting listed under slightly different names (e.g. "Junior Engineer (JE) Examination
// 2026" vs "SSC Junior Engineer (Civil/Electrical/Mechanical) 2026") must not duplicate —
// but different streams of the same programme must stay separate.
function sameJob(a, b) {
  if (a.id === b.id) return true;
  const sameOrg = a.source === b.source || overlap(tokens(a.org), tokens(b.org)) >= 0.6;
  if (!sameOrg) return false;
  const ta = tokens(a.title), tb = tokens(b.title);
  if (disciplineConflict(ta, tb)) return false;
  return overlap(ta, tb) >= 0.7;
}

function mergeJobs(store, found, sourceId) {
  const now = new Date().toISOString();
  let added = 0, updated = 0;
  for (const raw of found) {
    if (!raw?.org || !raw?.title || raw.fit === "low") continue;
    for (const k of ["org", "title", "summary", "advtNo", "posts", "fitLabel"]) {
      if (typeof raw[k] === "string") raw[k] = raw[k].replace(/\s+/g, " ").trim();
    }
    const id = slug(`${raw.org}-${raw.title}`);
    const candidate = { id, source: sourceId, org: raw.org, title: raw.title };
    const existing = store.jobs.find(j => sameJob(j, candidate));
    const job = {
      id,
      source: sourceId,
      org: raw.org,
      title: raw.title,
      summary: raw.summary || "",
      advtNo: raw.advtNo || null,
      posts: raw.posts || null,
      url: raw.url,
      status: raw.status,
      opensOn: raw.opensOn || null,
      closesOn: raw.closesOn || null,
      expected: raw.expected || null,
      fit: raw.fit,
      fitLabel: raw.fitLabel,
      tags: Array.isArray(raw.tags) ? raw.tags.slice(0, 2) : []
    };
    if (existing) {
      const before = JSON.stringify(existing);
      // Newer scrape wins, but never let a null wipe out a known value, and keep
      // the original id/firstSeen so links and NEW badges stay stable.
      for (const [k, v] of Object.entries(job)) {
        if (k === "id") continue;
        if (v !== null && v !== undefined && !(Array.isArray(v) && v.length === 0)) existing[k] = v;
      }
      if (JSON.stringify(existing) !== before) updated++;
    } else {
      store.jobs.push({ ...job, firstSeen: now });
      added++;
      console.log(`  NEW: ${job.org} — ${job.title} (${job.status}, closes ${job.closesOn || "?"})`);
    }
  }
  return { added, updated };
}

function dedupeStore(store) {
  const kept = [];
  for (const job of store.jobs) {
    const twin = kept.find(j => sameJob(j, job));
    if (!twin) { kept.push(job); continue; }
    // Merge into the earlier entry: filled fields win, earliest firstSeen kept.
    for (const [k, v] of Object.entries(job)) {
      if (k === "id" || k === "firstSeen") continue;
      if (v !== null && v !== undefined && !(Array.isArray(v) && v.length === 0)) twin[k] = v;
    }
    if (job.firstSeen < twin.firstSeen) twin.firstSeen = job.firstSeen;
    console.log(`  deduped: ${job.id} -> ${twin.id}`);
  }
  store.jobs = kept;
}

function autoClose(store) {
  const today = todayISO();
  for (const job of store.jobs) {
    if (job.status === "open" && job.closesOn && job.closesOn < today) {
      job.status = "closed";
    }
  }
  // Drop closed jobs older than 45 days to keep the file lean.
  const cutoff = new Date(Date.now() - 45 * 86_400_000).toISOString().slice(0, 10);
  store.jobs = store.jobs.filter(j => j.status !== "closed" || !j.closesOn || j.closesOn >= cutoff);
}

async function main() {
  loadEnvFile();
  if (!process.env.GEMINI_API_KEY) {
    console.error("GEMINI_API_KEY is not set (env var or .env file).");
    process.exit(1);
  }

  const args = process.argv.slice(2);
  const force = args.includes("--force");
  const onlyIdx = args.indexOf("--only");
  const only = onlyIdx !== -1 ? args[onlyIdx + 1] : null;

  const sources = loadJSON(join(ROOT, "config", "sources.json"), []);
  const state = loadJSON(STATE_FILE, { lastRun: {}, hashes: {} });
  const store = loadJSON(JOBS_FILE, { lastScan: null, jobs: [] });

  let geminiCalls = 0, totalAdded = 0, totalUpdated = 0;
  const now = Date.now();

  for (const source of sources) {
    if (only && source.id !== only) continue;
    const last = state.lastRun[source.id] || 0;
    const dueAt = last + source.cadenceMinutes * 60_000;
    if (!force && now < dueAt) continue;
    if (geminiCalls >= MAX_GEMINI_CALLS_PER_RUN) {
      console.log(`Budget: ${MAX_GEMINI_CALLS_PER_RUN} Gemini calls reached, deferring remaining sources.`);
      break;
    }

    console.log(`[${source.id}] checking (${source.method})...`);
    try {
      let found = null;
      if (source.method === "grounded") {
        geminiCalls++;
        found = await extractGrounded(source);
      } else {
        const html = await fetchPage(source.url);
        const text = htmlToText(html);
        const hash = sha256(text);
        if (!force && state.hashes[source.id] === hash) {
          console.log("  unchanged, skipping Gemini");
          state.lastRun[source.id] = now;
          continue;
        }
        geminiCalls++;
        found = await extractFromContent(source, text);
        state.hashes[source.id] = hash;
      }
      const { added, updated } = mergeJobs(store, found, source.id);
      totalAdded += added; totalUpdated += updated;
      state.lastRun[source.id] = now;
      console.log(`  ${found.length} postings parsed, ${added} new, ${updated} updated`);
    } catch (err) {
      console.error(`  FAILED: ${err.message}`);
      // Leave lastRun untouched so the source retries next cycle.
    }
  }

  dedupeStore(store);
  autoClose(store);
  store.lastScan = new Date().toISOString();
  store.jobs.sort((a, b) => (a.closesOn || "9999") < (b.closesOn || "9999") ? -1 : 1);

  writeFileSync(JOBS_FILE, JSON.stringify(store, null, 2) + "\n");
  writeFileSync(STATE_FILE, JSON.stringify(state, null, 2) + "\n");
  console.log(`Done. ${geminiCalls} Gemini calls, ${totalAdded} new jobs, ${totalUpdated} updated. ${store.jobs.length} tracked.`);
}

main().catch(err => { console.error(err); process.exit(1); });
