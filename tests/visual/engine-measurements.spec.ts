import { type Page, test } from "@playwright/test";
/**
 * Typing, open and save times on both editing engines, report-only.
 *
 * For each document, engine (the default one and the canonical session) and
 * host (React, Vue), on one page and one machine:
 *
 * - open: from handing the editor the file's bytes to the first page painted,
 *   and to the last layout pass before layout goes quiet (fully laid out);
 * - typing: per key, from keydown to the end of the first task after the next
 *   animation frame, typed at a caret in the middle of the document;
 * - save: the host's `save()` in rounds of a text edit, then Enter, then
 *   another text edit, each followed by a save in the same session, with the
 *   path each save took (a selective text patch, a structural splice or the
 *   full repack), and a forced full repack for reference.
 *
 * Each test writes one JSON file into `FOLIO_MEASURE_OUT` (default
 * `.cache/engine-measurements`); `scripts/engine-measurements-report.ts` turns
 * the files into tables. It asserts no number: the numbers depend on the
 * machine. A test fails only when a step cannot run at all, after it has
 * written what it measured. Run it with
 * `bunx playwright test --project=engine-measurements --workers=1`.
 *
 * - `FOLIO_MEASURE_KEYS`: measured keys per document (default 300).
 * - `FOLIO_MEASURE_HOSTS`: `react`, `vue` or both, comma-separated.
 * - `FOLIO_MEASURE_DOCUMENTS`: document ids to run, comma-separated.
 * - `FOLIO_MEASURE_MODE`: `editing` (default) or `suggesting`.
 * - `FOLIO_MEASURE_PROFILE=1`: also write a CPU profile of a typing burst and
 *   of one more open for each test, which the numbers above never include.
 *
 * A second set types on the canonical session with term highlighting
 * (anonymization terms) off and on, on two long documents, React host only;
 * it skips, with a note in its JSON file, on a build without highlighting.
 */
import { mkdirSync, writeFileSync } from "node:fs";
import { cpus } from "node:os";
import { join } from "node:path";

import { FIXTURES } from "../../benchmarks/fixtures";
import { createDocx } from "../../packages/core/src/docx/rezip";
import { fromMarkdown } from "../../packages/core/src/markdown/fromMarkdown";
import { PLAYGROUND_HOSTS } from "../parity/playgroundHosts";
import {
  contract,
  numberedList,
  plainParagraphs,
  tables,
} from "../support/generatedDocuments";

const WARM_KEYS = 5;
const MEASURED_KEYS = Number(process.env["FOLIO_MEASURE_KEYS"] ?? 300);
const PROFILED_KEYS = 40;
const PAUSE_MS = 150;
const OPEN_SAMPLES = 3;
const SAVE_ROUNDS = 3;
const QUIET_MS = 1500;
const STEP_TIMEOUT_MS = 180_000;
const OUT_DIRECTORY =
  process.env["FOLIO_MEASURE_OUT"] ?? ".cache/engine-measurements";
const PROFILE = process.env["FOLIO_MEASURE_PROFILE"] === "1";
const MODE =
  process.env["FOLIO_MEASURE_MODE"] === "suggesting" ? "suggesting" : "editing";

const generated = (markdown: () => string) => async () =>
  new Uint8Array(await createDocx(fromMarkdown(markdown())));

/**
 * The benchmark corpus (small to large; the large one is a real document with
 * bookmarks and fields), a long contract, a long list and many tables.
 */
const DOCUMENTS = [
  ...FIXTURES.map((fixture) => ({
    id: fixture.size,
    label: `${fixture.size} benchmark fixture (${fixture.name})`,
    bytes: async () => new Uint8Array(fixture.buffer),
  })),
  {
    id: "contract",
    label: "contract, 65 articles (generated)",
    bytes: generated(() => contract(65)),
  },
  {
    id: "list",
    label: "numbered list, 3,500 items (generated)",
    bytes: generated(() => numberedList(3500)),
  },
  {
    id: "tables",
    label: "40 tables of 30 rows (generated)",
    bytes: generated(() => tables(40, 30)),
  },
];

const ENGINES = ["legacy", "canonical"] as const;
const HOSTS = [
  { name: "react", origin: "" },
  { name: "vue", origin: PLAYGROUND_HOSTS.vue },
] as const;

const selected = (variable: string): Set<string> | null => {
  const value = process.env[variable];
  return value ? new Set(value.split(",").map((item) => item.trim())) : null;
};
const hostFilter = selected("FOLIO_MEASURE_HOSTS");
const documentFilter = selected("FOLIO_MEASURE_DOCUMENTS");

type Recorder = {
  lastLayoutAt: number;
  lastCompleteAt: number;
  layoutCompletes: number;
  measuredBlocks: number;
  selectiveAttempts: { written: boolean; structural: boolean }[];
  frames: number[];
};

declare global {
  var __engineMeasure: Recorder | undefined;
}

/** Times every layout pass and save attempt through the instrumentation hook. */
const installRecorder = () => {
  const recorder: Recorder = {
    lastLayoutAt: 0,
    lastCompleteAt: 0,
    layoutCompletes: 0,
    measuredBlocks: 0,
    selectiveAttempts: [],
    frames: [],
  };
  globalThis.__engineMeasure = recorder;
  const completed = () => {
    recorder.layoutCompletes += 1;
    recorder.lastCompleteAt = performance.now();
    recorder.lastLayoutAt = recorder.lastCompleteAt;
  };
  globalThis.__folioLayoutInstrumentation = {
    onLayoutStart: () => {
      recorder.lastLayoutAt = performance.now();
    },
    onLayoutComplete: completed,
    onLayoutError: completed,
    onMeasureBlock: () => {
      recorder.measuredBlocks += 1;
    },
    onSelectiveSaveAttempt: ({ written, structural }) => {
      recorder.selectiveAttempts.push({ written, structural });
    },
  };
  // Keydown to the end of the first task after the next frame: the edit and
  // the layout and paint it schedules.
  window.addEventListener(
    "keydown",
    () => {
      const start = performance.now();
      requestAnimationFrame(() => {
        const channel = new MessageChannel();
        channel.port1.onmessage = () =>
          recorder.frames.push(performance.now() - start);
        channel.port2.postMessage(undefined);
      });
    },
    { capture: true },
  );
};

/** Wait until no layout pass has started or ended for `quietMs`. */
const waitForQuiet = (page: Page, quietMs: number) =>
  page.waitForFunction(
    (quiet) =>
      performance.now() - (globalThis.__engineMeasure?.lastLayoutAt ?? 0) >=
      quiet,
    quietMs,
    { timeout: STEP_TIMEOUT_MS, polling: 100 },
  );

type OpenSample = {
  firstPageMs: number | null;
  laidOutMs: number | null;
  loadResolvedMs: number;
  layouts: number;
  pages: number;
  pageElements: number;
};

/**
 * Hand the editor the bytes and time the first page and the full layout. The
 * first page counts as painted at the end of the first task after the first
 * frame whose first page shows text other than the document it replaced.
 */
const openOnce = (page: Page, base64: string) =>
  page.evaluate(
    async ({ data, quietMs, timeoutMs }): Promise<OpenSample> => {
      const ref = globalThis.__folioPlayground?.getEditorRef();
      const recorder = globalThis.__engineMeasure;
      if (!ref || !recorder)
        throw new TypeError("The playground editor is not ready.");
      const bytes = Uint8Array.from(atob(data), (char) => char.charCodeAt(0));
      const firstPageText = () =>
        document.querySelector(".layout-page-content")?.textContent?.trim() ??
        "";
      const before = firstPageText();
      const completesBefore = recorder.layoutCompletes;
      const start = performance.now();
      const deadline = start + timeoutMs;
      const firstPage = new Promise<number | null>((resolve) => {
        const frame = () => {
          const text = firstPageText();
          if (text !== "" && text !== before) {
            const channel = new MessageChannel();
            channel.port1.onmessage = () => resolve(performance.now());
            channel.port2.postMessage(undefined);
          } else if (performance.now() > deadline) resolve(null);
          else requestAnimationFrame(frame);
        };
        requestAnimationFrame(frame);
      });
      await ref.loadDocumentBuffer(bytes);
      const loadResolvedAt = performance.now();
      // The default engine creates its editor view on first use.
      ref.ensureEditorView();
      const firstPageAt = await firstPage;
      const laidOutAt = await new Promise<number | null>((resolve) => {
        const check = () => {
          const now = performance.now();
          if (
            recorder.layoutCompletes > completesBefore &&
            now - recorder.lastLayoutAt >= quietMs
          )
            resolve(recorder.lastCompleteAt);
          else if (now > deadline) resolve(null);
          else setTimeout(check, 50);
        };
        check();
      });
      return {
        firstPageMs: firstPageAt === null ? null : firstPageAt - start,
        laidOutMs: laidOutAt === null ? null : laidOutAt - start,
        loadResolvedMs: loadResolvedAt - start,
        layouts: recorder.layoutCompletes - completesBefore,
        pages: ref.getTotalPages(),
        pageElements: document.querySelectorAll(".layout-page").length,
      };
    },
    { data: base64, quietMs: QUIET_MS, timeoutMs: STEP_TIMEOUT_MS },
  );

/** A fresh page on the playground's own small document, laid out. */
const freshPage = async (page: Page, url: string) => {
  await page.goto(url);
  await page.waitForSelector(".layout-page");
  await waitForQuiet(page, QUIET_MS);
};

/**
 * Put the caret three characters into the middle paragraph with text, on a
 * painted page in view, as a person types.
 */
const caretInMiddle = async (page: Page) => {
  const paragraphs = await page.evaluate(() => {
    const ref = globalThis.__folioPlayground?.getEditorRef();
    const view = ref?.getEditor()?.getView();
    if (!ref || !view) throw new TypeError("No editor view.");
    const positions: number[] = [];
    view.state.doc.descendants((node, position) => {
      if (node.type.name === "paragraph" && node.textContent.length > 4)
        positions.push(position);
      return true;
    });
    const middle = positions[Math.floor(positions.length / 2)];
    if (middle === undefined) throw new TypeError("No paragraph to type in.");
    // A host may find a paragraph to scroll to only on a page it has painted.
    const paraId: unknown = view.state.doc.nodeAt(middle)?.attrs["paraId"];
    const pageNumber =
      typeof paraId === "string"
        ? ref
            .getEditor()
            ?.getPositionPage({ story: "main", blockId: paraId, offset: 3 })
        : null;
    if (pageNumber !== undefined && pageNumber !== null)
      ref.scrollToPage(pageNumber);
    return { middle, count: positions.length };
  });
  await page.waitForTimeout(1000);
  const scrolled = await page.evaluate((middle) => {
    const ref = globalThis.__folioPlayground?.getEditorRef();
    const editor = ref?.getEditor();
    const view = editor?.getView();
    if (!ref || !editor || !view) throw new TypeError("No editor view.");
    editor.setSelection(middle + 3, middle + 3);
    const paraId: unknown = view.state.doc.nodeAt(middle)?.attrs["paraId"];
    const inView = typeof paraId === "string" && ref.scrollToParaId(paraId);
    editor.focus();
    return inView;
  }, paragraphs.middle);
  return { paragraphs: paragraphs.count, scrolledToParagraph: scrolled };
};

const readRecorder = (page: Page) =>
  page.evaluate(() => {
    const recorder = globalThis.__engineMeasure;
    if (!recorder) throw new TypeError("No recorder.");
    return {
      frames: recorder.frames.length,
      layouts: recorder.layoutCompletes,
      measuredBlocks: recorder.measuredBlocks,
    };
  });

/** Type `count` keys at a person's pace. */
const typeKeys = async (page: Page, count: number) => {
  for (let key = 0; key < count; key += 1) {
    // oxlint-disable-next-line no-await-in-loop -- keys are typed one after another
    await page.keyboard.press("x");
    // oxlint-disable-next-line no-await-in-loop -- a person's pause between keys
    await page.waitForTimeout(PAUSE_MS);
  }
};

/** Per-key latency of the measured keys, typed in the middle paragraph. */
const measureTyping = async (page: Page) => {
  const caret = await caretInMiddle(page);
  await typeKeys(page, WARM_KEYS);
  await waitForQuiet(page, 500);
  const before = await readRecorder(page);
  await typeKeys(page, MEASURED_KEYS);
  await waitForQuiet(page, 500);
  const after = await readRecorder(page);
  const frames = await page.evaluate(
    (from) => globalThis.__engineMeasure?.frames.slice(from) ?? [],
    before.frames,
  );
  return {
    ...caret,
    keys: frames.length,
    frameMs: frames,
    layoutsPerKey: (after.layouts - before.layouts) / MEASURED_KEYS,
    measuredBlocksPerKey:
      (after.measuredBlocks - before.measuredBlocks) / MEASURED_KEYS,
  };
};

type SaveSample = {
  ms: number;
  path:
    | "selective text patch"
    | "structural splice"
    | "full repack: selective declined"
    | "full repack: selective not tried";
  bytes: number | null;
};

const saveOnce = (page: Page, selective: boolean) =>
  page.evaluate(async (preferSelective): Promise<SaveSample> => {
    const ref = globalThis.__folioPlayground?.getEditorRef();
    const recorder = globalThis.__engineMeasure;
    if (!ref || !recorder) throw new TypeError("No editor.");
    recorder.selectiveAttempts = [];
    const start = performance.now();
    const buffer = await ref.save(
      preferSelective ? undefined : { selective: false },
    );
    const ms = performance.now() - start;
    const attempt = recorder.selectiveAttempts.at(-1);
    return {
      ms,
      path:
        attempt === undefined
          ? "full repack: selective not tried"
          : !attempt.written
            ? "full repack: selective declined"
            : attempt.structural
              ? "structural splice"
              : "selective text patch",
      bytes: buffer?.byteLength ?? null,
    };
  }, selective);

const profiled = async <T>(
  page: Page,
  file: string | null,
  action: () => Promise<T>,
): Promise<T> => {
  if (file === null) return action();
  const cdp = await page.context().newCDPSession(page);
  await cdp.send("Profiler.enable");
  await cdp.send("Profiler.setSamplingInterval", { interval: 200 });
  await cdp.send("Profiler.start");
  const result = await action();
  const { profile } = await cdp.send("Profiler.stop");
  mkdirSync(join(OUT_DIRECTORY, "profiles"), { recursive: true });
  writeFileSync(join(OUT_DIRECTORY, "profiles", file), JSON.stringify(profile));
  await cdp.detach();
  return result;
};

test.use({ viewport: { width: 1280, height: 900 }, deviceScaleFactor: 1 });

// Document, then engine: both engines of one document run back to back on
// the same machine, so a slower stretch of the machine hits both.
for (const host of HOSTS) {
  if (hostFilter && !hostFilter.has(host.name)) continue;
  for (const measured of DOCUMENTS) {
    if (documentFilter && !documentFilter.has(measured.id)) continue;
    for (const engine of ENGINES) {
      test(`engine measurements: ${measured.label}, ${engine} engine, ${host.name} host`, async ({
        page,
      }) => {
        test.setTimeout(1_200_000);
        const bytes = await measured.bytes();
        const base64 = Buffer.from(bytes).toString("base64");
        const name = `${host.name}-${measured.id}-${engine}`;
        const errors: string[] = [];
        const result = {
          host: host.name,
          engine,
          document: {
            id: measured.id,
            label: measured.label,
            bytes: bytes.byteLength,
          },
          mode: MODE,
          runner: {
            cpu: cpus()[0]?.model ?? "unknown",
            cores: cpus().length,
          },
          open: [] as OpenSample[],
          typing: {
            keys: 0,
            frameMs: [] as number[],
            layoutsPerKey: 0,
            measuredBlocksPerKey: 0,
            paragraphs: 0,
            scrolledToParagraph: false,
          },
          save: {
            text: [] as SaveSample[],
            structural: [] as SaveSample[],
            textAfterStructural: [] as SaveSample[],
            full: [] as SaveSample[],
          },
          errors,
        };
        const step = async (label: string, run: () => Promise<void>) => {
          try {
            await run();
          } catch (error) {
            errors.push(
              `${label}: ${error instanceof Error ? error.message : String(error)}`,
            );
          }
        };
        const query = new URLSearchParams({ mode: MODE });
        if (engine === "canonical") query.set("session", "canonical");
        const url = `${host.origin}/?${query.toString()}`;
        await page.addInitScript(installRecorder);
        // Each open on a fresh page, as a person opens a file; the document
        // stays open for the typing and saving that follow.
        await step("open", async () => {
          for (let sample = 0; sample < OPEN_SAMPLES; sample += 1) {
            // oxlint-disable-next-line no-await-in-loop -- opens are timed one after another
            await freshPage(page, url);
            // oxlint-disable-next-line no-await-in-loop -- opens are timed one after another
            result.open.push(await openOnce(page, base64));
          }
        });
        if (errors.length === 0)
          await step("typing", async () => {
            Object.assign(result.typing, await measureTyping(page));
            await profiled(
              page,
              PROFILE ? `${name}-typing.cpuprofile` : null,
              () => typeKeys(page, PROFILED_KEYS),
            );
          });
        // Each round: a text edit, Enter, a text edit, each saved in the
        // same session, so later rounds save after earlier structural edits.
        const edits = [
          ["text", "y"],
          ["structural", "Enter"],
          ["textAfterStructural", "y"],
        ] as const;
        if (errors.length === 0)
          await step("save", async () => {
            for (let round = 0; round < SAVE_ROUNDS; round += 1) {
              for (const [kind, key] of edits) {
                // oxlint-disable-next-line no-await-in-loop -- one edit before each save
                await page.keyboard.press(key);
                // oxlint-disable-next-line no-await-in-loop -- save after layout settles
                await waitForQuiet(page, 500);
                // oxlint-disable-next-line no-await-in-loop -- saves are timed one after another
                result.save[kind].push(await saveOnce(page, true));
              }
            }
            await page.keyboard.press("z");
            await waitForQuiet(page, 500);
            result.save.full.push(await saveOnce(page, false));
          });
        if (PROFILE && errors.length === 0)
          await step("profiled open", async () => {
            await freshPage(page, url);
            await profiled(page, `${name}-open.cpuprofile`, () =>
              openOnce(page, base64),
            );
          });
        mkdirSync(OUT_DIRECTORY, { recursive: true });
        writeFileSync(
          join(OUT_DIRECTORY, `${name}.json`),
          `${JSON.stringify(result, null, 2)}\n`,
        );
        const missing = [
          ...result.open.flatMap((sample) =>
            sample.firstPageMs === null || sample.laidOutMs === null
              ? ["an open did not finish"]
              : [],
          ),
          ...[
            ...result.save.text,
            ...result.save.structural,
            ...result.save.textAfterStructural,
          ].flatMap((sample) =>
            sample.bytes === null ? ["a save returned null"] : [],
          ),
        ];
        if (errors.length > 0 || missing.length > 0)
          throw new Error(
            `${name}: ${[...errors, ...new Set(missing)].join("; ")}`,
          );
      });
    }
  }
}

const TERM_DOCUMENTS = [
  {
    id: "paragraphs",
    label: "plain paragraphs, 1,500 (generated)",
    bytes: generated(() => plainParagraphs(1500)),
  },
  {
    id: "contract-100",
    label: "contract, 100 articles (generated)",
    bytes: generated(() => contract(100)),
  },
];

/** Highlight the first word's terms; false when the build cannot. */
const highlightTerms = (page: Page) =>
  page.evaluate(async () => {
    const bridge: unknown = Reflect.get(globalThis, "__folioParity");
    const anonymize =
      typeof bridge === "object" && bridge !== null
        ? Reflect.get(bridge, "anonymizeFirstWord")
        : undefined;
    if (typeof anonymize !== "function") return false;
    await anonymize.call(bridge);
    return true;
  });

if (hostFilter === null || hostFilter.has("react"))
  for (const measured of TERM_DOCUMENTS) {
    if (documentFilter && !documentFilter.has(measured.id)) continue;
    for (const terms of ["off", "on"] as const) {
      test(`typing with term highlighting ${terms}: ${measured.label}, canonical engine, react host`, async ({
        page,
      }) => {
        test.setTimeout(900_000);
        const bytes = await measured.bytes();
        const query = new URLSearchParams({ mode: MODE, session: "canonical" });
        await page.addInitScript(installRecorder);
        await freshPage(page, `/?${query.toString()}`);
        const open = await openOnce(
          page,
          Buffer.from(bytes).toString("base64"),
        );
        let skipped: string | null = null;
        if (terms === "on") {
          if (await highlightTerms(page)) await waitForQuiet(page, QUIET_MS);
          else skipped = "this build has no term highlighting to turn on";
        }
        const typing = skipped === null ? await measureTyping(page) : null;
        mkdirSync(OUT_DIRECTORY, { recursive: true });
        writeFileSync(
          join(
            OUT_DIRECTORY,
            `react-${measured.id}-canonical-terms-${terms}.json`,
          ),
          `${JSON.stringify(
            {
              host: "react",
              engine: "canonical",
              variant: { terms },
              document: {
                id: measured.id,
                label: measured.label,
                bytes: bytes.byteLength,
              },
              mode: MODE,
              runner: {
                cpu: cpus()[0]?.model ?? "unknown",
                cores: cpus().length,
              },
              open: [open],
              typing,
              skipped,
            },
            null,
            2,
          )}\n`,
        );
        test.skip(skipped !== null, skipped ?? "");
      });
    }
  }
