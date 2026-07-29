// REAL measure benchmark: loads real .docx fixtures into the actual folio
// editor in real Chromium with real fonts, and measures the REAL
// CanvasRenderingContext2D.measureText cost (call count + cumulative wall-time
// spent inside font measurement) during a COLD full-document layout, A/B-ing
// the pretext SegmentFitEngine (flag ON, "/") against the legacy word-walk
// (flag OFF, "/?segmentfit=off" == what upstream folio ships).
//
// No fake canvas. No synthetic paragraphs. Each (fixture, engine) runs in a
// FRESH browser context => the folio measure caches start cold (module state is
// per-realm), which is exactly how production opens a document.
//
// The docx is delivered via the editor's own file-open path (#file-input ->
// file.arrayBuffer() -> parse -> layout), so nothing is stubbed. measureText is
// patched via addInitScript before any page script runs; counters are reset
// after the initial (tiny default-doc) mount, so the numbers reflect only the
// uploaded document's layout.
//
// Usage:
//   FIXTURES=/tmp/nb_500_clean.tsv OUT=... bun bench-measure-browser.mjs
//   BASE=http://localhost:4200 LIMIT=1 bun bench-measure-browser.mjs   # smoke
import { appendFileSync, readFileSync, writeFileSync } from "node:fs";

import { chromium } from "@playwright/test";

const BASE = process.env.BASE ?? "http://localhost:4200";
const FIXTURES = process.env.FIXTURES ?? "/tmp/nb_500_clean.tsv";
const LIMIT = process.env.LIMIT ? Number(process.env.LIMIT) : Infinity;
const OUT = process.env.OUT ?? "./measure-browser.jsonl";
const SUMMARY = process.env.SUMMARY ?? OUT.replace(/\.jsonl$/, ".summary.json");
const SETTLE_POLL_MS = 100;
const SETTLE_STABLE = 3; // consecutive equal page counts => layout settled
const LOAD_TIMEOUT_MS = Number(process.env.LOAD_TIMEOUT_MS ?? 60000);

const PATCH = () => {
  const proto = CanvasRenderingContext2D.prototype;
  const orig = proto.measureText;
  window.__mtN = 0;
  window.__mtMs = 0;
  proto.measureText = function patched(text) {
    const a = performance.now();
    const r = orig.call(this, text);
    window.__mtMs += performance.now() - a;
    window.__mtN += 1;
    return r;
  };
  window.__mtReset = () => {
    window.__mtN = 0;
    window.__mtMs = 0;
  };
};

const ENGINES = [
  ["pretext", "/"],
  ["legacy", "/?segmentfit=off"],
];

const loadFixture = async (browser, name, filePath, engineName, urlPath) => {
  const context = await browser.newContext({ viewport: { width: 1280, height: 900 } });
  await context.addInitScript(PATCH);
  const page = await context.newPage();
  const row = { name, filePath, engine: engineName };
  const errors = [];
  page.on("pageerror", (e) => errors.push(String(e).slice(0, 160)));
  try {
    await page.goto(`${BASE}${urlPath}`, { waitUntil: "load", timeout: LOAD_TIMEOUT_MS });
    // editor mounted + test hook installed + initial (default doc) layout done
    await page.waitForFunction(() => typeof globalThis.__folioPlayground?.getEditorRef === "function", null, {
      timeout: LOAD_TIMEOUT_MS,
    });
    await page.waitForSelector('[data-testid="folio-editor"]', { timeout: LOAD_TIMEOUT_MS }).catch(() => {});
    // confirm the flag state actually matches what we intend
    row.flagOn = await page.evaluate(() => globalThis.__folioFeatureFlags?.segmentFitLineBreaking === true);

    // reset the measureText counters AFTER the default-doc mount so we capture
    // only the uploaded document's cold layout.
    await page.evaluate(() => globalThis.__mtReset?.());
    const t0 = Date.now();
    await page.setInputFiles("#file-input", filePath);

    // wait for layout to settle: getTotalPages() > 0 and stable for N polls,
    // or a page-level parse error surfaces.
    let stable = 0;
    let last = -1;
    let pages = 0;
    const deadline = Date.now() + LOAD_TIMEOUT_MS;
    while (Date.now() < deadline) {
      pages = await page.evaluate(() => globalThis.__folioPlayground?.getEditorRef?.()?.getTotalPages?.() ?? 0);
      if (pages > 0 && pages === last) {
        stable += 1;
        if (stable >= SETTLE_STABLE) break;
      } else {
        stable = 0;
      }
      last = pages;
      await page.waitForTimeout(SETTLE_POLL_MS);
    }
    row.wallMs = Date.now() - t0;
    row.pages = pages;
    const mt = await page.evaluate(() => ({ n: globalThis.__mtN ?? 0, ms: globalThis.__mtMs ?? 0 }));
    row.measureTextCalls = mt.n;
    row.measureTextMs = Math.round(mt.ms * 100) / 100;
    row.outcome = pages > 0 ? "ok" : "no-pages";
  } catch (error) {
    row.outcome = "error";
    row.error = String(error).split("\n")[0].slice(0, 200);
  } finally {
    if (errors.length) row.pageErrors = errors.slice(0, 2);
    await context.close().catch(() => {});
  }
  return row;
};

// ---- stats ----------------------------------------------------------------
const quantile = (s, q) => {
  if (!s.length) return null;
  const p = (s.length - 1) * q;
  const lo = Math.floor(p);
  const hi = Math.ceil(p);
  return lo === hi ? s[lo] : s[lo] + (s[hi] - s[lo]) * (p - lo);
};
const stats = (vals) => {
  const n = vals.filter((v) => Number.isFinite(v)).sort((a, b) => a - b);
  if (!n.length) return null;
  const sum = n.reduce((a, b) => a + b, 0);
  return {
    n: n.length,
    sum: Math.round(sum),
    mean: Math.round((sum / n.length) * 100) / 100,
    p50: Math.round(quantile(n, 0.5) * 100) / 100,
    p95: Math.round(quantile(n, 0.95) * 100) / 100,
    p99: Math.round(quantile(n, 0.99) * 100) / 100,
    max: Math.round(n.at(-1) * 100) / 100,
  };
};

// ---- main -----------------------------------------------------------------
const fixtures = readFileSync(FIXTURES, "utf8")
  .trim()
  .split("\n")
  .map((line, i) => {
    const [sz, path] = line.split("\t");
    return { name: `nb_${String(i).padStart(4, "0")}`, size: Number(sz), path };
  })
  .slice(0, LIMIT);

writeFileSync(OUT, "");
const browser = await chromium.launch({ headless: true });
const rows = [];
let done = 0;

// Concurrency pool. The headline metric (measureText CALL COUNT) is
// deterministic per (doc, engine) and unaffected by CPU contention, so running
// CONC docs in parallel is safe for it; measureTextMs is wall-time and under
// load is only indicative (we treat call count as the real signal, and the
// paired call counts are what the conclusion rests on).
const CONC = Number(process.env.CONC ?? 5);
const oneDoc = async (fx) => {
  const perEngine = {};
  for (const [engineName, urlPath] of ENGINES) {
    const row = await loadFixture(browser, fx.name, fx.path, engineName, urlPath);
    row.size = fx.size;
    rows.push(row);
    appendFileSync(OUT, `${JSON.stringify(row)}\n`);
    perEngine[engineName] = row;
  }
  done += 1;
  const p = perEngine.pretext;
  const l = perEngine.legacy;
  console.error(
    `[${done}/${fixtures.length}] ${fx.name} (${fx.size}B) pages=${l?.pages ?? "?"} ` +
      `calls L=${l?.measureTextCalls ?? "?"} P=${p?.measureTextCalls ?? "?"}`,
  );
};

const queue = [...fixtures];
const worker = async () => {
  while (queue.length) {
    const fx = queue.shift();
    if (fx) await oneDoc(fx);
  }
};
await Promise.all(Array.from({ length: CONC }, () => worker()));
await browser.close();

// ---- aggregate ------------------------------------------------------------
const byEngine = {};
for (const eng of ["legacy", "pretext"]) {
  const er = rows.filter((r) => r.engine === eng && r.outcome === "ok");
  byEngine[eng] = {
    docs: er.length,
    measureTextCalls: stats(er.map((r) => r.measureTextCalls)),
    measureTextMs: stats(er.map((r) => r.measureTextMs)),
    pages: stats(er.map((r) => r.pages)),
  };
}
// paired per-doc deltas (docs where both engines produced pages)
const paired = [];
const byName = new Map();
for (const r of rows) byName.set(`${r.name}:${r.engine}`, r);
for (const fx of fixtures) {
  const l = byName.get(`${fx.name}:legacy`);
  const p = byName.get(`${fx.name}:pretext`);
  if (l?.outcome === "ok" && p?.outcome === "ok" && l.measureTextCalls > 0) {
    paired.push({
      name: fx.name,
      size: fx.size,
      callsL: l.measureTextCalls,
      callsP: p.measureTextCalls,
      callRedPct: 1 - p.measureTextCalls / l.measureTextCalls,
      mtMsL: l.measureTextMs,
      mtMsP: p.measureTextMs,
      pagesMatch: l.pages === p.pages,
    });
  }
}
const summary = {
  harness: "measure-browser",
  base: BASE,
  fixtures: fixtures.length,
  paired: paired.length,
  pageMismatches: paired.filter((x) => !x.pagesMatch).length,
  byEngine,
  totals: {
    legacyCalls: byEngine.legacy.measureTextCalls?.sum,
    pretextCalls: byEngine.pretext.measureTextCalls?.sum,
    callReductionPct:
      byEngine.legacy.measureTextCalls && byEngine.pretext.measureTextCalls
        ? Math.round((1 - byEngine.pretext.measureTextCalls.sum / byEngine.legacy.measureTextCalls.sum) * 1000) / 10
        : null,
    legacyMtMs: byEngine.legacy.measureTextMs?.sum,
    pretextMtMs: byEngine.pretext.measureTextMs?.sum,
    mtMsReductionPct:
      byEngine.legacy.measureTextMs && byEngine.pretext.measureTextMs
        ? Math.round((1 - byEngine.pretext.measureTextMs.sum / byEngine.legacy.measureTextMs.sum) * 1000) / 10
        : null,
  },
  perDocCallReduction: stats(paired.map((x) => x.callRedPct * 100)),
};
writeFileSync(SUMMARY, `${JSON.stringify(summary, null, 2)}\n`);
console.log(JSON.stringify(summary, null, 2));
console.error(`\nrows: ${OUT}\nsummary: ${SUMMARY}`);
