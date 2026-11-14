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
 * - save: the host's `save()` after a text edit and after a structural edit
 *   (Enter), with the save path each one took (selective or full repack), and
 *   a forced full repack for reference.
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
 */
import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { cpus } from "node:os";
import { join } from "node:path";

import { createDocx } from "../../packages/core/src/docx/rezip";
import { fromMarkdown } from "../../packages/core/src/markdown/fromMarkdown";
import { PLAYGROUND_HOSTS } from "../parity/playgroundHosts";
import { contract, numberedList } from "../support/generatedDocuments";

const WARM_KEYS = 5;
const MEASURED_KEYS = Number(process.env["FOLIO_MEASURE_KEYS"] ?? 300);
const PROFILED_KEYS = 40;
const PAUSE_MS = 150;
const OPEN_SAMPLES = 3;
const SAVE_SAMPLES = 3;
const QUIET_MS = 1500;
const STEP_TIMEOUT_MS = 180_000;
const OUT_DIRECTORY =
  process.env["FOLIO_MEASURE_OUT"] ?? ".cache/engine-measurements";
const PROFILE = process.env["FOLIO_MEASURE_PROFILE"] === "1";
const MODE =
  process.env["FOLIO_MEASURE_MODE"] === "suggesting" ? "suggesting" : "editing";

const fixture = (name: string) => () =>
  new Uint8Array(readFileSync(join(import.meta.dirname, "fixtures", name)));
const generated = (markdown: () => string) => async () =>
  new Uint8Array(await createDocx(fromMarkdown(markdown())));

/** Small, real, long contract, long list, long table. */
const DOCUMENTS = [
  {
    id: "small",
    label: "small document (docx-editor-demo.docx)",
    bytes: fixture("docx-editor-demo.docx"),
  },
  {
    id: "real",
    label: "real document (podily-bps.docx)",
    bytes: fixture("podily-bps.docx"),
  },
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
    id: "table",
    label: "long table (performance-long-split-table.docx)",
    bytes: fixture("performance-long-split-table.docx"),
  },
] as const;

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
  selectiveAttempts: boolean[];
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
    onSelectiveSaveAttempt: ({ written }) => {
      recorder.selectiveAttempts.push(written);
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
};

/** Hand the editor the bytes and time the first page and the full layout. */
const openOnce = (page: Page, base64: string) =>
  page.evaluate(
    async ({ data, quietMs, timeoutMs }): Promise<OpenSample> => {
      const ref = globalThis.__folioPlayground?.getEditorRef();
      const recorder = globalThis.__engineMeasure;
      if (!ref || !recorder)
        throw new TypeError("The playground editor is not ready.");
      const bytes = Uint8Array.from(atob(data), (char) => char.charCodeAt(0));
      const completesBefore = recorder.layoutCompletes;
      const afterNextFrame = (resolve: (at: number) => void) =>
        requestAnimationFrame(() => {
          const channel = new MessageChannel();
          channel.port1.onmessage = () => resolve(performance.now());
          channel.port2.postMessage(undefined);
        });
      // The first text the painter puts on a page after the bytes arrive.
      let observer: MutationObserver | undefined;
      const firstPage = new Promise<number>((resolve) => {
        const onPage = (node: Node) =>
          (node instanceof Element ? node : node.parentElement)?.closest(
            ".layout-page-content",
          ) !== null && (node.textContent ?? "").trim() !== "";
        observer = new MutationObserver((records) => {
          for (const record of records) {
            const nodes =
              record.type === "characterData"
                ? [record.target]
                : [...record.addedNodes];
            if (nodes.some(onPage)) {
              observer?.disconnect();
              afterNextFrame(resolve);
              return;
            }
          }
        });
        observer.observe(document.body, {
          childList: true,
          subtree: true,
          characterData: true,
        });
      });
      const start = performance.now();
      await ref.loadDocumentBuffer(bytes);
      const loadResolvedAt = performance.now();
      // The default engine creates its editor view on first use.
      ref.ensureEditorView();
      const deadline = start + timeoutMs;
      const firstPageAt = await Promise.race([
        firstPage,
        new Promise<null>((resolve) =>
          setTimeout(() => resolve(null), timeoutMs),
        ),
      ]);
      observer?.disconnect();
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
      };
    },
    { data: base64, quietMs: QUIET_MS, timeoutMs: STEP_TIMEOUT_MS },
  );

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
    const pageNumber = ref.getEditorRef()?.getPageNumberForPmPos(middle + 3);
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

type SaveSample = {
  ms: number;
  path: "selective" | "full: selective refused" | "full: selective not tried";
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
          ? "full: selective not tried"
          : attempt
            ? "selective"
            : "full: selective refused",
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
        await page.addInitScript(installRecorder);
        await page.goto(`${host.origin}/?${query.toString()}`);
        await page.waitForSelector(".layout-page");
        await waitForQuiet(page, QUIET_MS);
        await step("open", async () => {
          for (let sample = 0; sample < OPEN_SAMPLES; sample += 1) {
            // oxlint-disable-next-line no-await-in-loop -- opens are timed one after another
            result.open.push(await openOnce(page, base64));
          }
        });
        if (errors.length === 0)
          await step("typing", async () => {
            Object.assign(result.typing, await caretInMiddle(page));
            for (let key = 0; key < WARM_KEYS; key += 1) {
              // oxlint-disable-next-line no-await-in-loop -- keys are typed one after another
              await page.keyboard.press("x");
              // oxlint-disable-next-line no-await-in-loop -- a person's pause between keys
              await page.waitForTimeout(PAUSE_MS);
            }
            await waitForQuiet(page, 500);
            const before = await readRecorder(page);
            for (let key = 0; key < MEASURED_KEYS; key += 1) {
              // oxlint-disable-next-line no-await-in-loop -- keys are typed one after another
              await page.keyboard.press("x");
              // oxlint-disable-next-line no-await-in-loop -- a person's pause between keys
              await page.waitForTimeout(PAUSE_MS);
            }
            await waitForQuiet(page, 500);
            const after = await readRecorder(page);
            const frames = await page.evaluate(
              (from) => globalThis.__engineMeasure?.frames.slice(from) ?? [],
              before.frames,
            );
            result.typing.keys = frames.length;
            result.typing.frameMs = frames;
            result.typing.layoutsPerKey =
              (after.layouts - before.layouts) / MEASURED_KEYS;
            result.typing.measuredBlocksPerKey =
              (after.measuredBlocks - before.measuredBlocks) / MEASURED_KEYS;
            await profiled(
              page,
              PROFILE ? `${name}-typing.cpuprofile` : null,
              async () => {
                for (let key = 0; key < PROFILED_KEYS; key += 1) {
                  // oxlint-disable-next-line no-await-in-loop -- keys are typed one after another
                  await page.keyboard.press("x");
                  // oxlint-disable-next-line no-await-in-loop -- a person's pause between keys
                  await page.waitForTimeout(PAUSE_MS);
                }
              },
            );
          });
        if (errors.length === 0)
          await step("save after a text edit", async () => {
            for (let sample = 0; sample < SAVE_SAMPLES; sample += 1) {
              if (sample > 0) {
                // oxlint-disable-next-line no-await-in-loop -- one edit before each save
                await page.keyboard.press("y");
              }
              // oxlint-disable-next-line no-await-in-loop -- save after layout settles
              await waitForQuiet(page, 500);
              // oxlint-disable-next-line no-await-in-loop -- saves are timed one after another
              result.save.text.push(await saveOnce(page, true));
            }
          });
        if (errors.length === 0)
          await step("save after Enter", async () => {
            for (let sample = 0; sample < SAVE_SAMPLES; sample += 1) {
              // oxlint-disable-next-line no-await-in-loop -- one edit before each save
              await page.keyboard.press("Enter");
              // oxlint-disable-next-line no-await-in-loop -- save after layout settles
              await waitForQuiet(page, 500);
              // oxlint-disable-next-line no-await-in-loop -- saves are timed one after another
              result.save.structural.push(await saveOnce(page, true));
            }
            await page.keyboard.press("z");
            await waitForQuiet(page, 500);
            result.save.full.push(await saveOnce(page, false));
          });
        if (PROFILE && errors.length === 0)
          await step("profiled open", async () => {
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
          ...[...result.save.text, ...result.save.structural].flatMap(
            (sample) => (sample.bytes === null ? ["a save returned null"] : []),
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
