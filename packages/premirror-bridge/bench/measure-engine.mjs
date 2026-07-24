// Comprehensive micro-benchmark for folio's paragraph MEASURE hot path
// (`measureParagraph`), A/B-ing the pretext SegmentFitEngine (flag ON) against
// the legacy word-walk (flag OFF == what upstream folio ships).
//
// WHY this harness and not the browser ones:
//   - bench-redline3-stats.mjs measures the whole redline3 app end-to-end.
//   - tests/perf/segmentfit-baseline.mjs measures WARM full-document relayout,
//     where its own note says pretext is expected to be at parity (folio's
//     width cache already covers repeat measurement).
//   The pretext win is COLD-cache / first-pass / changed-text measurement, and
//   the honest, machine-independent proxy for real font-measurement cost is the
//   number of canvas `measureText` calls (real canvas measureText is the
//   expensive op; pretext's whole purpose is to not call it). This harness
//   drives the real hot path headlessly through the deterministic fake canvas
//   (which counts calls AND whose per-call cost scales with string length, so
//   wall-time is a faithful relative signal too).
//
// Metrics per (archetype x width x engine):
//   - coldCalls   : mean canvas measureText calls for a FRESH paragraph
//                   (unique text every rep == "you just edited this line")
//   - coldMs      : wall-time distribution (p50/p95/p99/mean/sd) of one cold
//                   measure, warmup dropped
//   - warmCalls   : calls when re-measuring the SAME text (cache steady state)
//   - opsPerSec   : 1000 / mean(coldMs)
// Plus a document-scale macro: measure a K-paragraph document cold, total ms
// and total canvas calls (the "open / full relayout" cost that users feel).
//
// Modes:
//   MODE=full   (default) : run legacy AND pretext (needs @stll/premirror-bridge)
//   MODE=legacy           : run legacy only (works on an upstream checkout that
//                           has no segment-fit engine at all)
//
// Usage:
//   bun bench-measure-engine.mjs
//   MODE=legacy LABEL=upstream bun bench-measure-engine.mjs
//   M=400 WARMUP=25 WIDTHS=50,120,300,600 DOC_PARAS=500 bun bench-measure-engine.mjs
import { writeFileSync } from "node:fs";

// ---- config ---------------------------------------------------------------
const MODE = process.env.MODE ?? "full";
const LABEL = process.env.LABEL ?? (MODE === "legacy" ? "legacy-only" : "ours");
const M = Number(process.env.M ?? 200); // cold measure reps per cell
const WARMUP = Number(process.env.WARMUP ?? 15);
const WARM_REPS = Number(process.env.WARM_REPS ?? 200);
const WIDTHS = (process.env.WIDTHS ?? "50,120,300,600").split(",").map(Number);
const DOC_PARAS = Number(process.env.DOC_PARAS ?? 500);
const DOC_REPS = Number(process.env.DOC_REPS ?? 5);
const DOC_WIDTH = Number(process.env.DOC_WIDTH ?? 600);
const OUT = process.env.OUT ?? `./measure-engine-${LABEL}.json`;

// ---- dynamic imports (legacy mode never touches segment-fit / bridge) -----
const measureMod = await import("@stll/folio-core/layout-engine/measure/measureParagraph");
const fakeMod = await import(
  "@stll/folio-core/layout-engine/measure/__tests__/fakeTextMeasure"
);
const flagsMod = await import("@stll/folio-core/layout-engine/measure/featureFlags");
const cacheMod = await import("@stll/folio-core/layout-engine/measure/cache");
const { measureParagraph } = measureMod;
const { clearAllCaches } = cacheMod; // reset the per-word/per-slice width cache => true cold
const { withFakeTextMeasure, uppercaseAwareCharWidth } = fakeMod;
const { setFolioMeasurementFlags } = flagsMod;

let segmentMod = null;
let pretextEngine = null;
let clearPreparedCache = () => {};
if (MODE !== "legacy") {
  segmentMod = await import("@stll/folio-core/layout-engine/measure/segmentFit");
  const bridge = await import("@stll/premirror-bridge");
  pretextEngine = bridge.pretextSegmentFitEngine;
  clearPreparedCache = bridge.clearPreparedCache;
}

// ---- engine toggles -------------------------------------------------------
const useLegacy = () => {
  setFolioMeasurementFlags(undefined);
  segmentMod?.resetSegmentFitEngine();
  clearPreparedCache();
};
const usePretext = () => {
  setFolioMeasurementFlags({ segmentFitLineBreaking: true });
  segmentMod.setSegmentFitEngine(pretextEngine);
  clearPreparedCache();
};

const ENGINES =
  MODE === "legacy" ? [["legacy", useLegacy]] : [
    ["legacy", useLegacy],
    ["pretext", usePretext],
  ];

// ---- paragraph corpus (seed => unique-but-structurally-stable text) -------
const WORD_BANK =
  "the quick brown fox jumps over a lazy dog while parties hereto agree that any dispute arising under this agreement shall be resolved amicably".split(
    " ",
  );

// Novel prose: each position gets a realistic-length but DISTINCT token
// (index-suffixed), modelling the high novel-token density of real legal prose
// (party names, numbers, defined terms) rather than a 24-word loop the width
// cache would saturate. Coldness across reps comes from clearAllCaches(), so
// no per-rep seed is needed here.
const proseRuns = (wordCount) => {
  const words = [];
  for (let i = 0; i < wordCount; i += 1) words.push(`${WORD_BANK[i % WORD_BANK.length]}${i}`);
  return [{ kind: "text", text: words.join(" "), fontFamily: "Stub", fontSize: 12 }];
};

const CJK = "甲乙丙丁戊己庚辛壬癸子丑寅卯辰巳午未申酉戌亥";
const ARCHETYPES = {
  "prose-12w": () => proseRuns(12),
  "prose-100w": () => proseRuns(100),
  "prose-400w": () => proseRuns(400),
  "overlong-400c": () => [
    { kind: "text", text: "x".repeat(400), fontFamily: "Stub", fontSize: 12 },
  ],
  "cjk-120c": () => [
    { kind: "text", text: CJK.repeat(6).slice(0, 120), fontFamily: "Stub", fontSize: 12 },
  ],
  // 96 novel words split across 16 alternating font runs => cross-run word walking
  "mixed-16runs": () => {
    const runs = [];
    for (let i = 0; i < 16; i += 1) {
      const w = [];
      for (let j = 0; j < 6; j += 1) {
        const k = i * 6 + j;
        w.push(`${WORD_BANK[k % WORD_BANK.length]}${k}`);
      }
      runs.push({
        kind: "text",
        text: `${w.join(" ")} `,
        fontFamily: i % 2 === 0 ? "Stub" : "Stub2",
        fontSize: 12,
      });
    }
    return runs;
  },
};

const para = (runs) => ({ kind: "paragraph", id: "p", runs });

// ---- stats ----------------------------------------------------------------
const quantile = (sorted, q) => {
  if (sorted.length === 0) return null;
  const pos = (sorted.length - 1) * q;
  const lo = Math.floor(pos);
  const hi = Math.ceil(pos);
  return lo === hi ? sorted[lo] : sorted[lo] + (sorted[hi] - sorted[lo]) * (pos - lo);
};
const round = (v, d = 4) => (v == null ? null : Math.round(v * 10 ** d) / 10 ** d);
const stats = (values) => {
  const nums = values.filter((v) => Number.isFinite(v));
  if (nums.length === 0) return null;
  const mean = nums.reduce((a, b) => a + b, 0) / nums.length;
  const sd =
    nums.length > 1
      ? Math.sqrt(nums.reduce((a, b) => a + (b - mean) ** 2, 0) / (nums.length - 1))
      : 0;
  const s = [...nums].sort((a, b) => a - b);
  return {
    n: nums.length,
    mean: round(mean),
    sd: round(sd),
    p50: round(quantile(s, 0.5)),
    p95: round(quantile(s, 0.95)),
    p99: round(quantile(s, 0.99)),
    min: round(s.at(0)),
    max: round(s.at(-1)),
  };
};

// Reset ALL measurement state so the next measure is genuinely cold, for BOTH
// engines (folio's per-word/slice width+font caches AND pretext's prepared
// cache). No-op for the pretext cache in legacy mode.
const coldReset = () => {
  clearAllCaches();
  clearPreparedCache();
};

// ---- one cell: cold + warm for a given archetype/width --------------------
const runCell = (getCount, makeRuns, width) => {
  const block = para(makeRuns()); // identical text; coldness comes from coldReset()

  // Warmup (JIT) — dropped.
  for (let i = 0; i < WARMUP; i += 1) {
    coldReset();
    measureParagraph(block, width);
  }

  // Cold: cache cleared before each rep => true first-paint cost. Time only the
  // measure; count canvas calls per rep.
  const coldMs = [];
  let coldCallsTotal = 0;
  for (let i = 0; i < M; i += 1) {
    coldReset();
    const c0 = getCount();
    const t0 = performance.now();
    measureParagraph(block, width);
    coldMs.push(performance.now() - t0);
    coldCallsTotal += getCount() - c0;
  }

  // Warm: same text re-measured with caches intact (cache steady state).
  coldReset();
  measureParagraph(block, width); // prime
  const w0 = getCount();
  const warmMs = [];
  for (let i = 0; i < WARM_REPS; i += 1) {
    const t0 = performance.now();
    measureParagraph(block, width);
    warmMs.push(performance.now() - t0);
  }
  const warmCallsTotal = getCount() - w0;

  const cold = stats(coldMs);
  return {
    width,
    coldCallsPerMeasure: round(coldCallsTotal / M, 2),
    coldMs: cold,
    opsPerSec: cold && cold.mean > 0 ? Math.round(1000 / cold.mean) : null,
    warmCallsPerMeasure: round(warmCallsTotal / WARM_REPS, 2),
    warmMs: stats(warmMs),
  };
};

// ---- document-scale macro: cold "open document" of NOVEL content ----------
// Each paragraph gets genuinely novel tokens (uid-suffixed), modelling a
// document with no repeated content — the realistic high-water mark for a cold
// full-document measure pass. Cache is cleared once per rep (document open),
// then builds naturally across paragraphs within the pass.
const docParaKind = (uid) => {
  if (uid % 25 === 0) return "overlong";
  if (uid % 5 === 0) return 400;
  return 100;
};
const docParaRuns = (uid) => {
  const kind = docParaKind(uid);
  if (kind === "overlong") {
    return [
      { kind: "text", text: `${"x".repeat(380)}${uid}`, fontFamily: "Stub", fontSize: 12 },
    ];
  }
  const words = [];
  for (let i = 0; i < kind; i += 1) words.push(`${WORD_BANK[i % WORD_BANK.length]}${uid}_${i}`);
  return [{ kind: "text", text: words.join(" "), fontFamily: "Stub", fontSize: 12 }];
};

const runDoc = (getCount) => {
  // warmup pass (JIT), then measured reps with a fresh novel doc each time.
  for (let i = 0; i < DOC_PARAS; i += 1) {
    coldReset();
    measureParagraph(para(docParaRuns(9_000_000 + i)), DOC_WIDTH);
  }

  const ms = [];
  let callsTotal = 0;
  for (let r = 0; r < DOC_REPS; r += 1) {
    const rBlocks = [];
    for (let i = 0; i < DOC_PARAS; i += 1) {
      rBlocks.push(para(docParaRuns(r * DOC_PARAS + i)));
    }
    coldReset(); // cold document open; cache builds across the pass
    const c0 = getCount();
    const t0 = performance.now();
    for (const b of rBlocks) measureParagraph(b, DOC_WIDTH);
    ms.push(performance.now() - t0);
    callsTotal += getCount() - c0;
  }
  return {
    paragraphs: DOC_PARAS,
    width: DOC_WIDTH,
    reps: DOC_REPS,
    totalMs: stats(ms),
    canvasCallsPerDoc: Math.round(callsTotal / DOC_REPS),
  };
};

// ---- main -----------------------------------------------------------------
const result = {
  harness: "measure-engine",
  label: LABEL,
  mode: MODE,
  config: { M, WARMUP, WARM_REPS, WIDTHS, DOC_PARAS, DOC_REPS, DOC_WIDTH },
  engines: {},
};

withFakeTextMeasure((getCount) => {
  for (const [engineName, setEngine] of ENGINES) {
    const cells = {};
    for (const [archName, makeRuns] of Object.entries(ARCHETYPES)) {
      cells[archName] = {};
      for (const width of WIDTHS) {
        setEngine(); // reset caches + engine before each cell
        cells[archName][`w${width}`] = runCell(getCount, makeRuns, width);
      }
    }
    setEngine();
    const doc = runDoc(getCount);
    result.engines[engineName] = { cells, doc };
    console.error(`[${LABEL}] ${engineName}: doc ${doc.paragraphs}p -> ${doc.totalMs.p50}ms p50, ${doc.canvasCallsPerDoc} canvas calls`);
  }
}, { charWidth: uppercaseAwareCharWidth });

writeFileSync(OUT, `${JSON.stringify(result, null, 2)}\n`);
console.error(`written: ${OUT}`);

// ---- console summary table ------------------------------------------------
const pct = (from, to) => (from > 0 ? `${Math.round((1 - to / from) * 100)}%` : "—");
if (result.engines.legacy && result.engines.pretext) {
  console.log(`\n=== ${LABEL}: pretext vs legacy — COLD canvas measureText calls per paragraph ===`);
  console.log("archetype".padEnd(16), "width".padStart(6), "legacy".padStart(9), "pretext".padStart(9), "saved".padStart(7));
  for (const arch of Object.keys(ARCHETYPES)) {
    for (const width of WIDTHS) {
      const L = result.engines.legacy.cells[arch][`w${width}`];
      const P = result.engines.pretext.cells[arch][`w${width}`];
      console.log(
        arch.padEnd(16),
        String(width).padStart(6),
        String(L.coldCallsPerMeasure).padStart(9),
        String(P.coldCallsPerMeasure).padStart(9),
        pct(L.coldCallsPerMeasure, P.coldCallsPerMeasure).padStart(7),
      );
    }
  }
  const dl = result.engines.legacy.doc;
  const dp = result.engines.pretext.doc;
  console.log(`\n=== document-scale (${dl.paragraphs} paras @ ${dl.width}px, cold) ===`);
  console.log(`legacy : ${dl.totalMs.p50}ms p50 (${dl.totalMs.mean}±${dl.totalMs.sd}), ${dl.canvasCallsPerDoc} canvas calls`);
  console.log(`pretext: ${dp.totalMs.p50}ms p50 (${dp.totalMs.mean}±${dp.totalMs.sd}), ${dp.canvasCallsPerDoc} canvas calls`);
  console.log(`saved  : ${pct(dl.totalMs.p50, dp.totalMs.p50)} time, ${pct(dl.canvasCallsPerDoc, dp.canvasCallsPerDoc)} canvas calls`);
}
