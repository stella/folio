import { chromium, type Browser, type BrowserContext, type Page } from "@playwright/test";
import { TaggedError } from "better-result";
import { spawn } from "node:child_process";
import { mkdir, writeFile } from "node:fs/promises";
import { arch, cpus, loadavg, platform, release } from "node:os";
import { dirname, resolve as resolvePath } from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

import type {
  DocumentLoadPhase,
  HiddenEditorPhase,
  HiddenEditorStateReason,
  LayoutInstrumentation,
  LayoutPhase,
  LayoutRunReason,
} from "../src/layout-engine/layoutInstrumentation";

type CounterBucket = {
  count: number;
  totalMs: number;
};

type CacheState = "cold" | "warm";
type InstrumentationMode = "detailed" | "standard";
type PerfSuite = "baseline" | "corpus" | "smoke";
type ScenarioKind = "edit" | "open";

type PerfMilestone =
  | "bytes-available"
  | "complete-layout"
  | "docx-parsed"
  | "editor-state-ready"
  | "first-usable-page"
  | "flow-blocks-ready"
  | "fonts-ready"
  | "layout-start"
  | "measurement-ready"
  | "pagination-ready"
  | "prosemirror-ready"
  | "render-pages-ready"
  | "visible-pages-painted";

type PerfStats = {
  createElement: CounterBucket;
  documentLoadPhases: Record<DocumentLoadPhase, CounterBucket>;
  elements: number;
  getBoundingClientRect: CounterBucket;
  hiddenEditorPhases: Record<HiddenEditorPhase, CounterBucket>;
  hiddenPmElements: number;
  hiddenStateCreations: Record<HiddenEditorStateReason, number>;
  layoutCompletions: number;
  layoutErrors: { message: string; reason: LayoutRunReason }[];
  layoutPhases: Record<LayoutPhase, CounterBucket>;
  layoutReasons: Record<LayoutRunReason, number>;
  layoutStarts: Record<LayoutRunReason, number>;
  longTasks: {
    count: number;
    durationsMs: number[];
    maxMs: number;
    totalMs: number;
  };
  measureBlockCalls: number;
  measureText: CounterBucket;
  measuredBlockRange: { first: number; last: number } | null;
  milestones: Partial<Record<PerfMilestone, number>>;
  pages: number;
  renderedPages: number;
  uniqueMeasuredBlocks: number;
  visiblePageElements: number;
};

type BrowserMetrics = {
  JSHeapUsedSize?: number;
  LayoutDuration?: number;
  RecalcStyleDuration?: number;
  ScriptDuration?: number;
  TaskDuration?: number;
};

type PerfSample = {
  browser: {
    heapUsedMb: number;
    layoutMs: number;
    scriptMs: number;
    styleMs: number;
    taskMs: number;
  };
  iteration: number;
  stats: PerfStats;
  wallMs: number;
};

type Distribution = {
  max: number;
  median: number;
  min: number;
  p95: number;
};

type PerfScenarioResult = {
  aggregates: Record<string, Distribution>;
  cacheState: CacheState;
  kind: ScenarioKind;
  name: string;
  samples: PerfSample[];
};

type PerfReport = {
  environment: {
    architecture: string;
    browser: string;
    browserVersion: string;
    bunVersion: string;
    commit: string;
    cpu: string;
    deviceScaleFactor: number;
    instrumentation: InstrumentationMode;
    loadAverage: {
      after: number[];
      before: number[];
    };
    logicalCpuCount: number;
    nodeVersion: string;
    operatingSystem: string;
    operatingSystemRelease: string;
    repetitions: number;
    suite: PerfSuite;
    viewport: { height: number; width: number };
  };
  generatedAt: string;
  scenarios: PerfScenarioResult[];
  schemaVersion: 1;
  startedAt: string;
};

type CDPSession = Awaited<ReturnType<BrowserContext["newCDPSession"]>>;

type OpenScenario = {
  minimumPages: number;
  name: string;
  path: string;
};

type BrowserPerfStats = PerfStats & {
  measuredBlockIndexes: Set<number>;
  sampleStartedAt: number;
  trackAutomaticPagePaint: boolean;
};

class FolioPerfError extends TaggedError("FolioPerfError")<{
  message: string;
  cause?: unknown;
}>() {}

const HOST = "127.0.0.1";
const DEFAULT_PORT = 4201;
const DEVICE_SCALE_FACTOR = 2;
const VIEWPORT = { height: 900, width: 1280 } as const;
const DEFAULT_REPETITIONS = 3;
const EDIT_PARAGRAPH_COUNT = 1500;
const port = Number(process.env["FOLIO_PERF_PORT"] ?? DEFAULT_PORT);
const baseUrl = process.env["FOLIO_PERF_URL"] ?? `http://${HOST}:${port}`;
const skipBuild = process.env["FOLIO_PERF_SKIP_BUILD"] === "1";
const repetitions = readPositiveInteger(process.env["FOLIO_PERF_REPETITIONS"], DEFAULT_REPETITIONS);
const instrumentation: InstrumentationMode =
  process.env["FOLIO_PERF_DIAGNOSTICS"] === "1" ? "detailed" : "standard";
const suite = readSuite(process.env["FOLIO_PERF_SUITE"]);
const outputPath = resolvePath(
  rootDir(),
  process.env["FOLIO_PERF_OUTPUT"] ?? ".cache/performance/folio-perf.json",
);
const scenarioCatalog = {
  demo: {
    minimumPages: 1,
    name: "open docx-editor-demo.docx",
    path: "/?file=docx-editor-demo.docx",
  },
  longTable: {
    minimumPages: 1,
    name: "open generated long split table DOCX",
    path: "/?file=performance-long-split-table.docx",
  },
  mixedScript: {
    minimumPages: 1,
    name: "open generated mixed-script embedded-font DOCX",
    path: "/?file=performance-mixed-script-embedded-font.docx",
  },
  podily: {
    minimumPages: 1,
    name: "open podily-bps.docx",
    path: "/?file=podily-bps.docx",
  },
  sample: { minimumPages: 1, name: "open sample.docx", path: "/?file=sample.docx" },
  paragraphs: {
    minimumPages: 20,
    name: "open generated 1500-paragraph DOCX",
    path: "/?file=performance-1500-paragraphs.docx",
  },
} as const satisfies Record<string, OpenScenario>;
const openScenariosBySuite: Record<PerfSuite, readonly OpenScenario[]> = {
  baseline: [scenarioCatalog.demo, scenarioCatalog.podily, scenarioCatalog.paragraphs],
  corpus: [
    scenarioCatalog.sample,
    scenarioCatalog.demo,
    scenarioCatalog.podily,
    scenarioCatalog.paragraphs,
    scenarioCatalog.longTable,
    scenarioCatalog.mixedScript,
  ],
  smoke: [scenarioCatalog.demo],
};
const openScenarios = openScenariosBySuite[suite];

declare global {
  var __folioPerfCounters: BrowserPerfStats | undefined;
  var __recordFolioPerfMilestone: ((name: PerfMilestone) => void) | undefined;
  var __resetFolioPerfCounters: ((mode?: ScenarioKind) => void) | undefined;
  var __folioLayoutInstrumentation: LayoutInstrumentation | undefined;
  var __folioPlayground:
    | {
        getEditorRef: () => {
          getEditorRef: () => { getView: () => unknown | null } | null;
        } | null;
      }
    | undefined;
}

async function main(): Promise<void> {
  let previewProcess: ReturnType<typeof spawn> | null = null;
  let browser: Browser | null = null;
  const startedAt = new Date().toISOString();
  const loadAverageBefore = loadavg().map(round);

  try {
    if (!process.env["FOLIO_PERF_URL"]) {
      if (!skipBuild) {
        process.stderr.write("Building the production playground...\n");
        await runCommand({
          args: ["--filter", "@stll/playground", "build"],
          command: "bun",
          cwd: rootDir(),
        });
      }

      previewProcess = spawn(
        "bun",
        ["--filter", "@stll/playground", "preview", "--host", HOST, "--port", String(port)],
        {
          cwd: rootDir(),
          stdio: ["ignore", "pipe", "pipe"],
        },
      );
      previewProcess.stdout?.pipe(process.stderr);
      previewProcess.stderr?.pipe(process.stderr);
      await waitForServer(baseUrl);
    }

    browser = await chromium.launch({ headless: true });
    const scenarios = await profileScenarios(browser);
    const report: PerfReport = {
      environment: {
        architecture: arch(),
        browser: "chromium",
        browserVersion: browser.version(),
        bunVersion: process.versions.bun ?? "unavailable",
        commit: await readCommandOutput({
          args: ["rev-parse", "HEAD"],
          command: "git",
          cwd: rootDir(),
        }),
        cpu: cpus().at(0)?.model ?? "unknown",
        deviceScaleFactor: DEVICE_SCALE_FACTOR,
        instrumentation,
        loadAverage: {
          after: loadavg().map(round),
          before: loadAverageBefore,
        },
        logicalCpuCount: cpus().length,
        nodeVersion: process.version,
        operatingSystem: platform(),
        operatingSystemRelease: release(),
        repetitions,
        suite,
        viewport: VIEWPORT,
      },
      generatedAt: new Date().toISOString(),
      scenarios,
      schemaVersion: 1,
      startedAt,
    };
    const json = `${JSON.stringify(report, null, 2)}\n`;
    await mkdir(dirname(outputPath), { recursive: true });
    await writeFile(outputPath, json, "utf8");
    process.stderr.write(`Wrote machine-readable results to ${outputPath}\n`);
    process.stdout.write(json);
  } finally {
    await browser?.close();
    previewProcess?.kill();
  }
}

function rootDir(): string {
  let dir = fileURLToPath(new URL("../../..", import.meta.url));
  while (dir.endsWith("/") || dir.endsWith("\\")) {
    dir = dir.slice(0, -1);
  }
  return dir;
}

type CommandOptions = {
  args: string[];
  command: string;
  cwd: string;
};

async function runCommand({ command, args, cwd }: CommandOptions): Promise<void> {
  await new Promise<void>((resolve, reject) => {
    const child = spawn(command, args, {
      cwd,
      stdio: ["ignore", "pipe", "pipe"],
    });
    child.stdout?.pipe(process.stderr);
    child.stderr?.pipe(process.stderr);
    child.on("error", reject);
    child.on("exit", (code) => {
      if (code === 0) {
        resolve();
        return;
      }
      reject(
        new FolioPerfError({
          message: `${command} ${args.join(" ")} exited with ${code}`,
        }),
      );
    });
  });
}

function readCommandOutput({ command, args, cwd }: CommandOptions): Promise<string> {
  return new Promise<string>((resolve, reject) => {
    const child = spawn(command, args, { cwd, stdio: ["ignore", "pipe", "pipe"] });
    let stdout = "";
    let stderr = "";
    child.stdout?.on("data", (chunk: Buffer) => {
      stdout += chunk.toString();
    });
    child.stderr?.on("data", (chunk: Buffer) => {
      stderr += chunk.toString();
    });
    child.on("error", reject);
    child.on("exit", (code) => {
      if (code === 0) {
        resolve(stdout.trim());
        return;
      }
      reject(
        new FolioPerfError({
          message: `${command} ${args.join(" ")} exited with ${code}: ${stderr.trim()}`,
        }),
      );
    });
  });
}

async function waitForServer(url: string): Promise<void> {
  const startedAt = performance.now();
  while (performance.now() - startedAt < 20_000) {
    try {
      // oxlint-disable-next-line no-await-in-loop -- readiness probes must be sequential
      const response = await fetch(url);
      if (response.ok) {
        return;
      }
    } catch {
      // The preview process is still starting.
    }
    // oxlint-disable-next-line no-await-in-loop -- fixed back-off between readiness probes
    await new Promise<void>((resolve) => {
      setTimeout(resolve, 100);
    });
  }
  throw new FolioPerfError({ message: `Timed out waiting for ${url}` });
}

async function profileScenarios(browser: Browser): Promise<PerfScenarioResult[]> {
  const results: PerfScenarioResult[] = [];

  for (const scenario of openScenarios) {
    for (const cacheState of ["cold", "warm"] as const) {
      process.stderr.write(`Measuring ${scenario.name} (${cacheState})...\n`);
      // oxlint-disable-next-line no-await-in-loop -- benchmark scenarios run sequentially to avoid contention
      results.push(await profileOpenScenario({ browser, cacheState, scenario }));
    }
  }

  process.stderr.write("Measuring single-key editing latency (warm)...\n");
  results.push(await profileTypingScenario(browser));
  return results;
}

type ProfileOpenScenarioOptions = {
  browser: Browser;
  cacheState: CacheState;
  scenario: OpenScenario;
};

async function profileOpenScenario({
  browser,
  cacheState,
  scenario,
}: ProfileOpenScenarioOptions): Promise<PerfScenarioResult> {
  const samples: PerfSample[] = [];
  let warmContext: BrowserContext | null = null;

  try {
    if (cacheState === "warm") {
      warmContext = await createContext(browser);
      await primeContext(warmContext, scenario.path);
    }

    for (let iteration = 1; iteration <= repetitions; iteration += 1) {
      const context = warmContext ?? (await createContext(browser));
      try {
        // oxlint-disable-next-line no-await-in-loop -- samples run sequentially to avoid benchmark contention
        samples.push(await profileLoad({ context, iteration, scenario }));
      } finally {
        if (!warmContext) {
          // oxlint-disable-next-line no-await-in-loop -- close each cold context before the next sample
          await context.close();
        }
      }
    }
  } finally {
    await warmContext?.close();
  }

  return {
    aggregates: aggregateSamples(samples, "open"),
    cacheState,
    kind: "open",
    name: scenario.name,
    samples,
  };
}

async function profileTypingScenario(browser: Browser): Promise<PerfScenarioResult> {
  const context = await createContext(browser);
  const samples: PerfSample[] = [];

  try {
    await primeContext(context, `/?paragraphs=${EDIT_PARAGRAPH_COUNT}`);
    for (let iteration = 1; iteration <= repetitions; iteration += 1) {
      const page = await context.newPage();
      const cdp = await context.newCDPSession(page);
      await cdp.send("Performance.enable");
      try {
        // oxlint-disable-next-line no-await-in-loop -- editing samples run sequentially to avoid contention
        await page.goto(`${baseUrl}/?paragraphs=${EDIT_PARAGRAPH_COUNT}`);
        // oxlint-disable-next-line no-await-in-loop -- setup must finish before timing the edit
        await waitForEditor(page, 20);
        // oxlint-disable-next-line no-await-in-loop -- click establishes a stable caret before timing
        await page
          .locator(".layout-page")
          .first()
          .click({ position: { x: 120, y: 120 } });
        // oxlint-disable-next-line no-await-in-loop -- flush click-driven paint before resetting action counters
        await waitForTwoFrames(page);
        // oxlint-disable-next-line no-await-in-loop -- capture the visible content before the measured edit
        const visibleTextBefore = await page.locator(".layout-page-content").first().textContent();
        // oxlint-disable-next-line no-await-in-loop -- metrics bracket the measured action
        const before = await readMetrics(cdp);
        // oxlint-disable-next-line no-await-in-loop -- reset browser counters immediately before the action
        await page.evaluate(() => globalThis.__resetFolioPerfCounters?.("edit"));
        const startedAt = performance.now();
        // oxlint-disable-next-line no-await-in-loop -- one user-like key event is the measured action
        await page.keyboard.press("a");
        // oxlint-disable-next-line no-await-in-loop -- wait for the edited content to reach the visible page
        await page.waitForFunction(
          (previousText) => {
            const pageContent = document.querySelector(".layout-page-content");
            return pageContent !== null && pageContent.textContent !== previousText;
          },
          visibleTextBefore,
          { timeout: 20_000 },
        );
        // oxlint-disable-next-line no-await-in-loop -- include the paint opportunity after visible content changes
        await waitForTwoFrames(page);
        // oxlint-disable-next-line no-await-in-loop -- record the browser-clock paint boundary for this sample
        await page.evaluate(() => {
          globalThis.__recordFolioPerfMilestone?.("visible-pages-painted");
        });
        // oxlint-disable-next-line no-await-in-loop -- wait for transaction layout and paint
        await page.waitForFunction(
          () =>
            (globalThis.__folioPerfCounters?.layoutReasons.transaction ?? 0) > 0 &&
            globalThis.__folioPerfCounters?.milestones["complete-layout"] !== undefined,
          undefined,
          { timeout: 20_000 },
        );
        // oxlint-disable-next-line no-await-in-loop -- include the next paint opportunity
        await waitForTwoFrames(page);
        const wallMs = performance.now() - startedAt;
        // oxlint-disable-next-line no-await-in-loop -- metrics bracket the measured action
        const after = await readMetrics(cdp);
        // oxlint-disable-next-line no-await-in-loop -- read one completed sample before continuing
        samples.push({
          browser: metricDelta(before, after),
          iteration,
          stats: await readStats(page),
          wallMs: round(wallMs),
        });
      } finally {
        // oxlint-disable-next-line no-await-in-loop -- close each sample page before continuing
        await page.close();
      }
    }
  } finally {
    await context.close();
  }

  return {
    aggregates: aggregateSamples(samples, "edit"),
    cacheState: "warm",
    kind: "edit",
    name: "single key to transaction layout and paint",
    samples,
  };
}

async function createContext(browser: Browser): Promise<BrowserContext> {
  const context = await browser.newContext({
    deviceScaleFactor: DEVICE_SCALE_FACTOR,
    viewport: VIEWPORT,
  });
  await installCounters(context, instrumentation === "detailed");
  return context;
}

async function primeContext(context: BrowserContext, path: string): Promise<void> {
  const page = await context.newPage();
  try {
    await page.goto(`${baseUrl}${path}`);
    await waitForEditor(page, 1);
  } finally {
    await page.close();
  }
}

type ProfileLoadOptions = {
  context: BrowserContext;
  iteration: number;
  scenario: OpenScenario;
};

async function profileLoad({
  context,
  iteration,
  scenario,
}: ProfileLoadOptions): Promise<PerfSample> {
  const page = await context.newPage();
  const cdp = await context.newCDPSession(page);
  await cdp.send("Performance.enable");

  try {
    const before = await readMetrics(cdp);
    const startedAt = performance.now();
    await page.goto(`${baseUrl}${scenario.path}`);
    await waitForEditor(page, scenario.minimumPages);
    const wallMs = performance.now() - startedAt;
    const after = await readMetrics(cdp);
    return {
      browser: metricDelta(before, after),
      iteration,
      stats: await readStats(page),
      wallMs: round(wallMs),
    };
  } finally {
    await page.close();
  }
}

async function readMetrics(cdp: CDPSession): Promise<BrowserMetrics> {
  const result = await cdp.send("Performance.getMetrics");
  return Object.fromEntries(result.metrics.map((metric) => [metric.name, metric.value]));
}

function metricDelta(before: BrowserMetrics, after: BrowserMetrics): PerfSample["browser"] {
  return {
    heapUsedMb: round((after.JSHeapUsedSize ?? 0) / 1024 / 1024),
    layoutMs: round(((after.LayoutDuration ?? 0) - (before.LayoutDuration ?? 0)) * 1000),
    scriptMs: round(((after.ScriptDuration ?? 0) - (before.ScriptDuration ?? 0)) * 1000),
    styleMs: round(((after.RecalcStyleDuration ?? 0) - (before.RecalcStyleDuration ?? 0)) * 1000),
    taskMs: round(((after.TaskDuration ?? 0) - (before.TaskDuration ?? 0)) * 1000),
  };
}

async function waitForEditor(page: Page, minimumPages: number): Promise<void> {
  try {
    await page.waitForFunction(
      (pageCount) => {
        const perf = globalThis.__folioPerfCounters;
        return (
          document.querySelectorAll(".layout-page").length >= pageCount &&
          perf?.milestones["first-usable-page"] !== undefined &&
          perf.milestones["complete-layout"] !== undefined
        );
      },
      minimumPages,
      { timeout: 30_000 },
    );
    await waitForTwoFrames(page);
  } catch (error) {
    const diagnostics = await page.evaluate(() => ({
      errors: globalThis.__folioPerfCounters?.layoutErrors ?? [],
      milestones: globalThis.__folioPerfCounters?.milestones ?? {},
      pages: document.querySelectorAll(".layout-page").length,
      renderedPages: document.querySelectorAll(".layout-page-content").length,
    }));
    throw new FolioPerfError({
      cause: error,
      message: `Timed out waiting for a usable ${minimumPages}-page editor: ${JSON.stringify(diagnostics)}`,
    });
  }
}

async function waitForTwoFrames(page: Page): Promise<void> {
  await page.evaluate(
    () =>
      new Promise<void>((resolve) => {
        requestAnimationFrame(() => {
          requestAnimationFrame(() => resolve());
        });
      }),
  );
}

async function readStats(page: Page): Promise<PerfStats> {
  const stats = await page.evaluate(() => {
    const perf = globalThis.__folioPerfCounters;
    if (!perf) {
      return null;
    }
    return {
      createElement: roundBucket(perf.createElement),
      documentLoadPhases: { "docx-parse": roundBucket(perf.documentLoadPhases["docx-parse"]) },
      elements: document.querySelectorAll("*").length,
      getBoundingClientRect: roundBucket(perf.getBoundingClientRect),
      hiddenEditorPhases: {
        "editor-state": roundBucket(perf.hiddenEditorPhases["editor-state"]),
        "editor-view": roundBucket(perf.hiddenEditorPhases["editor-view"]),
        "to-prose-doc": roundBucket(perf.hiddenEditorPhases["to-prose-doc"]),
        "update-state": roundBucket(perf.hiddenEditorPhases["update-state"]),
      },
      hiddenPmElements: document.querySelectorAll(".paged-editor__hidden-pm-wrapper *").length,
      hiddenStateCreations: perf.hiddenStateCreations,
      layoutCompletions: perf.layoutCompletions,
      layoutErrors: perf.layoutErrors,
      layoutPhases: {
        "flow-blocks": roundBucket(perf.layoutPhases["flow-blocks"]),
        "header-footer": roundBucket(perf.layoutPhases["header-footer"]),
        "initial-fonts": roundBucket(perf.layoutPhases["initial-fonts"]),
        "layout-document": roundBucket(perf.layoutPhases["layout-document"]),
        "measure-blocks": roundBucket(perf.layoutPhases["measure-blocks"]),
        "render-pages": roundBucket(perf.layoutPhases["render-pages"]),
      },
      layoutReasons: perf.layoutReasons,
      layoutStarts: perf.layoutStarts,
      longTasks: {
        count: perf.longTasks.count,
        durationsMs: perf.longTasks.durationsMs.map(roundInBrowser),
        maxMs: roundInBrowser(perf.longTasks.maxMs),
        totalMs: roundInBrowser(perf.longTasks.totalMs),
      },
      measureBlockCalls: perf.measureBlockCalls,
      measureText: roundBucket(perf.measureText),
      measuredBlockRange: perf.measuredBlockRange,
      milestones: Object.fromEntries(
        Object.entries(perf.milestones).map(([name, value]) => [name, roundInBrowser(value)]),
      ),
      pages: document.querySelectorAll(".layout-page").length,
      renderedPages: document.querySelectorAll(".layout-page-content").length,
      uniqueMeasuredBlocks: perf.measuredBlockIndexes.size,
      visiblePageElements: document.querySelectorAll(".paged-editor__pages *").length,
    };

    function roundInBrowser(value: number): number {
      return Number(value.toFixed(2));
    }

    function roundBucket(bucket: CounterBucket): CounterBucket {
      return { count: bucket.count, totalMs: roundInBrowser(bucket.totalMs) };
    }
  });

  if (!stats) {
    throw new FolioPerfError({ message: "Folio performance counters were not installed" });
  }
  return stats;
}

async function installCounters(context: BrowserContext, detailed: boolean): Promise<void> {
  await context.addInitScript((enableDetailedCounters) => {
    const makeBucket = (): CounterBucket => ({ count: 0, totalMs: 0 });
    const makeLayoutPhaseCounters = (): Record<LayoutPhase, CounterBucket> => ({
      "flow-blocks": makeBucket(),
      "header-footer": makeBucket(),
      "initial-fonts": makeBucket(),
      "layout-document": makeBucket(),
      "measure-blocks": makeBucket(),
      "render-pages": makeBucket(),
    });
    const makeHiddenEditorPhaseCounters = (): Record<HiddenEditorPhase, CounterBucket> => ({
      "editor-state": makeBucket(),
      "editor-view": makeBucket(),
      "to-prose-doc": makeBucket(),
      "update-state": makeBucket(),
    });
    const makeHiddenStateCreationCounters = (): Record<HiddenEditorStateReason, number> => ({
      "external-document": 0,
      mount: 0,
    });
    const makeLayoutReasonCounters = (): Record<LayoutRunReason, number> => ({
      "font-ready": 0,
      initial: 0,
      "layout-input": 0,
      manual: 0,
      transaction: 0,
    });
    const makeCounters = (): BrowserPerfStats => ({
      createElement: makeBucket(),
      documentLoadPhases: { "docx-parse": makeBucket() },
      elements: 0,
      getBoundingClientRect: makeBucket(),
      hiddenEditorPhases: makeHiddenEditorPhaseCounters(),
      hiddenPmElements: 0,
      hiddenStateCreations: makeHiddenStateCreationCounters(),
      layoutCompletions: 0,
      layoutErrors: [],
      layoutPhases: makeLayoutPhaseCounters(),
      layoutReasons: makeLayoutReasonCounters(),
      layoutStarts: makeLayoutReasonCounters(),
      longTasks: { count: 0, durationsMs: [], maxMs: 0, totalMs: 0 },
      measureBlockCalls: 0,
      measureText: makeBucket(),
      measuredBlockIndexes: new Set<number>(),
      measuredBlockRange: null,
      milestones: {},
      pages: 0,
      renderedPages: 0,
      sampleStartedAt: performance.now(),
      trackAutomaticPagePaint: true,
      uniqueMeasuredBlocks: 0,
      visiblePageElements: 0,
    });
    const recordMilestone = (name: PerfMilestone, explicitTime?: number): void => {
      const perf = globalThis.__folioPerfCounters;
      if (!perf || perf.milestones[name] !== undefined) {
        return;
      }
      perf.milestones[name] = (explicitTime ?? performance.now()) - perf.sampleStartedAt;
    };

    globalThis.__folioPerfCounters = makeCounters();
    globalThis.__recordFolioPerfMilestone = recordMilestone;
    globalThis.__resetFolioPerfCounters = (mode = "open") => {
      const counters = makeCounters();
      counters.trackAutomaticPagePaint = mode === "open";
      globalThis.__folioPerfCounters = counters;
    };
    globalThis.__folioLayoutInstrumentation = {
      onDocumentLoadPhase(event) {
        const perf = globalThis.__folioPerfCounters;
        if (!perf) {
          return;
        }
        addToBucket(perf.documentLoadPhases[event.phase], event.durationMs);
        recordMilestone("docx-parsed");
      },
      onHiddenEditorPhase(event) {
        const perf = globalThis.__folioPerfCounters;
        if (!perf) {
          return;
        }
        addToBucket(perf.hiddenEditorPhases[event.phase], event.durationMs);
        if (event.phase === "to-prose-doc") {
          recordMilestone("prosemirror-ready");
        }
        if (event.phase === "editor-state") {
          recordMilestone("editor-state-ready");
        }
      },
      onHiddenEditorStateCreate(event) {
        const perf = globalThis.__folioPerfCounters;
        if (perf) {
          perf.hiddenStateCreations[event.reason] += 1;
        }
      },
      onLayoutComplete(event) {
        const perf = globalThis.__folioPerfCounters;
        if (!perf) {
          return;
        }
        perf.layoutCompletions += 1;
        perf.layoutReasons[event.reason] += 1;
        recordMilestone("complete-layout");
      },
      onLayoutError(event) {
        globalThis.__folioPerfCounters?.layoutErrors.push(event);
      },
      onLayoutStart(event) {
        const perf = globalThis.__folioPerfCounters;
        if (!perf) {
          return;
        }
        perf.layoutStarts[event.reason] += 1;
        recordMilestone("layout-start");
      },
      onLayoutPhase(event) {
        const perf = globalThis.__folioPerfCounters;
        if (!perf) {
          return;
        }
        addToBucket(perf.layoutPhases[event.phase], event.durationMs);
        const milestoneByPhase: Partial<Record<LayoutPhase, PerfMilestone>> = {
          "flow-blocks": "flow-blocks-ready",
          "initial-fonts": "fonts-ready",
          "layout-document": "pagination-ready",
          "measure-blocks": "measurement-ready",
          "render-pages": "render-pages-ready",
        };
        const milestone = milestoneByPhase[event.phase];
        if (milestone) {
          recordMilestone(milestone);
        }
      },
      onMeasureBlock(event) {
        const perf = globalThis.__folioPerfCounters;
        if (!perf) {
          return;
        }
        perf.measureBlockCalls += 1;
        perf.measuredBlockIndexes.add(event.blockIndex);
        const range = perf.measuredBlockRange;
        perf.measuredBlockRange = range
          ? {
              first: Math.min(range.first, event.blockIndex),
              last: Math.max(range.last, event.blockIndex),
            }
          : { first: event.blockIndex, last: event.blockIndex };
      },
    };

    const observeUsablePage = (): void => {
      let frameScheduled = false;
      const check = (): void => {
        if (!globalThis.__folioPerfCounters?.trackAutomaticPagePaint) {
          return;
        }
        const hasContent = document.querySelector(".layout-page-content [data-pm-start]") !== null;
        if (!hasContent || frameScheduled) {
          return;
        }
        frameScheduled = true;
        requestAnimationFrame(() => {
          requestAnimationFrame(() => {
            frameScheduled = false;
            if (document.querySelector(".layout-page-content [data-pm-start]") === null) {
              return;
            }
            recordMilestone("visible-pages-painted");
            const editor = globalThis.__folioPlayground?.getEditorRef()?.getEditorRef();
            if (editor) {
              recordMilestone("first-usable-page");
              return;
            }
            setTimeout(check, 0);
          });
        });
      };
      const observer = new MutationObserver(check);
      observer.observe(document, { childList: true, subtree: true });
      check();
    };
    if (document.readyState === "loading") {
      document.addEventListener("DOMContentLoaded", observeUsablePage, { once: true });
    } else {
      observeUsablePage();
    }

    try {
      new PerformanceObserver((list) => {
        for (const entry of list.getEntries()) {
          if (
            entry.entryType === "resource" &&
            entry.name.includes("/fixtures/") &&
            "responseEnd" in entry &&
            typeof entry.responseEnd === "number"
          ) {
            recordMilestone("bytes-available", entry.responseEnd);
          }
          if (entry.entryType !== "longtask") {
            continue;
          }
          const perf = globalThis.__folioPerfCounters;
          if (!perf || entry.duration <= 50) {
            continue;
          }
          perf.longTasks.count += 1;
          perf.longTasks.durationsMs.push(entry.duration);
          perf.longTasks.totalMs += entry.duration;
          perf.longTasks.maxMs = Math.max(perf.longTasks.maxMs, entry.duration);
        }
      }).observe({ entryTypes: ["longtask", "resource"] });
    } catch {
      // These performance entry types are unavailable in some browser modes.
    }

    if (!enableDetailedCounters) {
      return;
    }
    wrapMethod(CanvasRenderingContext2D.prototype, "measureText", "measureText");
    wrapMethod(Element.prototype, "getBoundingClientRect", "getBoundingClientRect");
    wrapMethod(Document.prototype, "createElement", "createElement");

    function addToBucket(bucket: CounterBucket, durationMs: number): void {
      bucket.count += 1;
      bucket.totalMs += durationMs;
    }

    function wrapMethod(
      owner: object,
      methodName: "createElement" | "getBoundingClientRect" | "measureText",
      bucketName: "createElement" | "getBoundingClientRect" | "measureText",
    ): void {
      const descriptor = Object.getOwnPropertyDescriptor(owner, methodName);
      const original = descriptor?.value;
      if (typeof original !== "function") {
        return;
      }
      Object.defineProperty(owner, methodName, {
        ...descriptor,
        value(this: unknown, ...args: unknown[]) {
          const startedAt = performance.now();
          try {
            return Reflect.apply(original, this, args);
          } finally {
            const perf = globalThis.__folioPerfCounters;
            if (perf) {
              addToBucket(perf[bucketName], performance.now() - startedAt);
            }
          }
        },
      });
    }
  }, detailed);
}

function aggregateSamples(
  samples: readonly PerfSample[],
  scenarioKind: ScenarioKind,
): Record<string, Distribution> {
  const metrics: Record<string, number[]> = {};
  const add = (name: string, value: number | undefined): void => {
    if (value === undefined) {
      return;
    }
    (metrics[name] ??= []).push(value);
  };
  const addBucket = (name: string, bucket: CounterBucket): void => {
    if (bucket.count > 0) {
      add(name, bucket.totalMs);
    }
  };

  for (const sample of samples) {
    const parse = sample.stats.documentLoadPhases["docx-parse"];
    const proseMirror = sample.stats.hiddenEditorPhases["to-prose-doc"];
    const hiddenEditorPhaseTotal = sumBucketMs(sample.stats.hiddenEditorPhases);
    const layoutPhaseTotal = sumBucketMs(sample.stats.layoutPhases);
    const layoutPipelinePhaseTotal = sumBucketMs(sample.stats.layoutPhases, "initial-fonts");
    const accountedPhaseMs = parse.totalMs + hiddenEditorPhaseTotal + layoutPhaseTotal;
    add("wallMs", sample.wallMs);
    add("browser.heapUsedMb", sample.browser.heapUsedMb);
    add("browser.layoutMs", sample.browser.layoutMs);
    add("browser.scriptMs", sample.browser.scriptMs);
    add("browser.styleMs", sample.browser.styleMs);
    add("browser.taskMs", sample.browser.taskMs);
    addBucket("docxParseMs", parse);
    addBucket("proseMirrorConversionMs", proseMirror);
    addBucket("editorStateCreationMs", sample.stats.hiddenEditorPhases["editor-state"]);
    addBucket("editorViewCreationMs", sample.stats.hiddenEditorPhases["editor-view"]);
    addBucket("fontReadinessMs", sample.stats.layoutPhases["initial-fonts"]);
    addBucket("flowBlockGenerationMs", sample.stats.layoutPhases["flow-blocks"]);
    addBucket("measurementMs", sample.stats.layoutPhases["measure-blocks"]);
    addBucket("paginationMs", sample.stats.layoutPhases["layout-document"]);
    addBucket("renderPagesMs", sample.stats.layoutPhases["render-pages"]);
    if (scenarioKind === "open") {
      add("firstUsablePageMs", sample.stats.milestones["first-usable-page"]);
    } else {
      add("keyToVisibleUpdateMs", sample.stats.milestones["visible-pages-painted"]);
    }
    add("completeLayoutMs", sample.stats.milestones["complete-layout"]);
    add("bytesAvailableMs", sample.stats.milestones["bytes-available"]);
    add("visiblePagesPaintedMs", sample.stats.milestones["visible-pages-painted"]);
    const renderPagesReadyMs = sample.stats.milestones["render-pages-ready"];
    const visiblePagesPaintedMs = sample.stats.milestones["visible-pages-painted"];
    if (renderPagesReadyMs !== undefined && visiblePagesPaintedMs !== undefined) {
      add("pagePaintingMs", Math.max(0, visiblePagesPaintedMs - renderPagesReadyMs));
    }
    add("accountedPhaseMs", accountedPhaseMs);
    const layoutStartMs = sample.stats.milestones["layout-start"];
    add(scenarioKind === "edit" ? "inputToLayoutStartMs" : "layoutStartMs", layoutStartMs);
    const completeLayoutMs = sample.stats.milestones["complete-layout"];
    if (completeLayoutMs !== undefined) {
      add("unaccountedToCompleteMs", Math.max(0, completeLayoutMs - accountedPhaseMs));
      const bytesAvailableMs = sample.stats.milestones["bytes-available"];
      if (bytesAvailableMs !== undefined) {
        const pipelineAfterBytesMs = Math.max(0, completeLayoutMs - bytesAvailableMs);
        add("pipelineAfterBytesMs", pipelineAfterBytesMs);
        add("unaccountedAfterBytesMs", Math.max(0, pipelineAfterBytesMs - accountedPhaseMs));
      }
      if (layoutStartMs !== undefined) {
        const layoutPipelineMs = Math.max(0, completeLayoutMs - layoutStartMs);
        add("layoutPipelineMs", layoutPipelineMs);
        add(
          "unaccountedInsideLayoutPipelineMs",
          Math.max(0, layoutPipelineMs - layoutPipelinePhaseTotal),
        );
      }
    }
    add("longTaskTotalMs", sample.stats.longTasks.totalMs);
    add("blocksMeasured", sample.stats.measureBlockCalls);
    add("uniqueBlocksMeasured", sample.stats.uniqueMeasuredBlocks);
    add("totalPages", sample.stats.pages);
    add("renderedPages", sample.stats.renderedPages);
    add("domElements", sample.stats.elements);
  }

  return Object.fromEntries(
    Object.entries(metrics).map(([name, values]) => [name, calculateDistribution(values)]),
  );
}

function sumBucketMs<TName extends string>(
  buckets: Record<TName, CounterBucket>,
  excluded?: TName,
): number {
  let total = 0;
  for (const [name, bucket] of Object.entries(buckets)) {
    if (name !== excluded) {
      total += bucket.totalMs;
    }
  }
  return total;
}

function calculateDistribution(values: readonly number[]): Distribution {
  const sorted = values.toSorted((left, right) => left - right);
  const medianIndex = Math.floor(sorted.length / 2);
  const p95Index = Math.max(0, Math.ceil(sorted.length * 0.95) - 1);
  return {
    max: round(sorted.at(-1) ?? 0),
    median: round(sorted.at(medianIndex) ?? 0),
    min: round(sorted.at(0) ?? 0),
    p95: round(sorted.at(p95Index) ?? 0),
  };
}

function readPositiveInteger(value: string | undefined, fallback: number): number {
  const parsed = Number(value);
  return Number.isSafeInteger(parsed) && parsed > 0 ? parsed : fallback;
}

function readSuite(value: string | undefined): PerfSuite {
  if (value === "smoke" || value === "corpus") {
    return value;
  }
  return "baseline";
}

function round(value: number): number {
  return Number(value.toFixed(2));
}

await main();
