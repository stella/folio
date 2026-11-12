import { type Page, test } from "@playwright/test";
/**
 * Scratch measurement only (never merged): the canonical session against the
 * legacy engine in the React playground, per document and mode.
 *
 * Per run: openMs (loadDocumentBuffer to the first completed layout of the new
 * document), per key inputMs (Event Timing processing, entries of 16 ms and
 * more), taskMs (keydown to a message posted from keydown), frameMs (keydown
 * to the end of the first task after the next animation frame), and saveMs
 * (median of three saves after typing). FOLIO_ZZ_PROFILE=<dir> writes a CPU
 * profile of the measured keys instead of trusting the timings.
 */
import {
  appendFileSync,
  mkdirSync,
  readFileSync,
  writeFileSync,
} from "node:fs";
import { join } from "node:path";

import { createDocx } from "../../packages/core/src/docx/rezip";
import { fromMarkdown } from "../../packages/core/src/markdown/fromMarkdown";

const WARM_KEYS = 5;
const MEASURED_KEYS = Number(process.env["FOLIO_ZZ_KEYS"] ?? 30);
const PAUSE_MS = 150;
const SAVES = 3;
const ROUND = process.env["FOLIO_ZZ_ROUND"] ?? "0";
const OUT = process.env["FOLIO_ZZ_OUT"];
const PROFILE = process.env["FOLIO_ZZ_PROFILE"];
const ORDER = (process.env["FOLIO_ZZ_ORDER"] ?? "legacy,canonical").split(",");
const ONLY_DOCS = process.env["FOLIO_ZZ_DOCS"]?.split(",");

const WORDS =
  "the supplier shall deliver goods services under this agreement subject to terms conditions warranty liability indemnity confidential information party notice period".split(
    " ",
  );
const contract = (articles: number): string => {
  let seed = 7;
  const random = () => {
    seed = (seed * 1_103_515_245 + 12_345) % 2 ** 31;
    return seed / 2 ** 31;
  };
  const sentence = (length: number) =>
    Array.from(
      { length },
      () => WORDS[Math.floor(random() * WORDS.length)] ?? "the",
    ).join(" ");
  const lines = ["# Services Agreement", ""];
  for (let article = 1; article <= articles; article += 1) {
    lines.push(`## ${article}. ${sentence(3)}`, "");
    for (let clause = 1; clause <= 8; clause += 1) {
      lines.push(
        `${article}.${clause} **${sentence(2)}** ${sentence(40)}. ${sentence(25)}.`,
        "",
      );
      if (clause % 3 === 0) {
        for (let item = 1; item <= 4; item += 1)
          lines.push(`${item}. ${sentence(18)};`);
        lines.push("");
      }
    }
    if (article % 10 === 0) {
      lines.push("| Item | Fee | Term |", "|---|---|---|");
      for (let row = 0; row < 12; row += 1)
        lines.push(`| ${sentence(3)} | ${row * 100} | ${sentence(2)} |`);
      lines.push("");
    }
  }
  return lines.join("\n");
};
const numberedList = (items: number): string =>
  Array.from({ length: items }, (_, index) => `1. Item ${index + 1}`).join(
    "\n",
  );

const DOCUMENTS: Record<string, () => Promise<number[]>> = {
  "list-3500": async () => [
    ...new Uint8Array(await createDocx(fromMarkdown(numberedList(3500)))),
  ],
  "contract-100": async () => [
    ...new Uint8Array(await createDocx(fromMarkdown(contract(100)))),
  ],
  lir011: async () => [
    ...readFileSync(process.env["FOLIO_ZZ_LIR011"] ?? "LIR011.docx"),
  ],
};

const median = (values: readonly number[]): number | null => {
  const sorted = values.toSorted((left, right) => left - right);
  return sorted.length === 0 ? null : (sorted[sorted.length >> 1] ?? null);
};

const open = (page: Page, bytes: number[]) =>
  page.evaluate(async (data) => {
    const ref = globalThis.__folioPlayground?.getEditorRef();
    if (!ref) throw new TypeError("The playground editor is not ready.");
    const layouts =
      (globalThis as { __zzLayouts?: number[] }).__zzLayouts ?? [];
    const started = performance.now();
    await ref.loadDocumentBuffer(new Uint8Array(data));
    const loaded = performance.now();
    ref.ensureEditorView();
    const deadline = started + 180_000;
    for (;;) {
      const view = ref.getEditor()?.getView();
      const count = view?.state.doc.childCount ?? 0;
      const laidOut = layouts.find((at) => at >= loaded);
      if (count > 500 && laidOut !== undefined)
        return {
          openMs: laidOut - started,
          loadMs: loaded - started,
          blocks: count,
        };
      if (performance.now() > deadline) throw new TypeError("Open timed out.");
      await new Promise((resolve) => setTimeout(resolve, 5));
    }
  }, bytes);

const prepare = (page: Page) =>
  page.evaluate(() => {
    const ref = globalThis.__folioPlayground?.getEditorRef();
    const editor = ref?.getEditor();
    const view = editor?.getView();
    if (!ref || !editor || !view) throw new TypeError("No editor view.");
    const paragraphs: number[] = [];
    view.state.doc.descendants((node, position) => {
      if (node.type.name === "paragraph" && node.textContent.length > 4)
        paragraphs.push(position);
      return true;
    });
    const middle = paragraphs[Math.floor(paragraphs.length / 2)];
    if (middle === undefined) throw new TypeError("No paragraph to type in.");
    const pageNumber = ref.getEditorRef()?.getPageNumberForPmPos(middle + 3);
    if (pageNumber !== undefined && pageNumber !== null)
      ref.scrollToPage(pageNumber);
    editor.setSelection(middle + 3, middle + 3);
    const paraId: unknown = view.state.doc.nodeAt(middle)?.attrs["paraId"];
    if (typeof paraId === "string") ref.scrollToParaId(paraId);
    editor.focus();
    const record = {
      input: new Map<number, number>(),
      tasks: [] as number[],
      frames: [] as number[],
    };
    new PerformanceObserver((list) => {
      for (const entry of list.getEntries() as PerformanceEventTiming[]) {
        if (entry.interactionId === 0) continue;
        record.input.set(
          entry.interactionId,
          (record.input.get(entry.interactionId) ?? 0) +
            entry.processingEnd -
            entry.processingStart,
        );
      }
    }).observe({ type: "event", durationThreshold: 16 });
    window.addEventListener(
      "keydown",
      () => {
        const start = performance.now();
        const task = new MessageChannel();
        task.port1.onmessage = () =>
          record.tasks.push(performance.now() - start);
        task.port2.postMessage(undefined);
        requestAnimationFrame(() => {
          const channel = new MessageChannel();
          channel.port1.onmessage = () =>
            record.frames.push(performance.now() - start);
          channel.port2.postMessage(undefined);
        });
      },
      { capture: true },
    );
    (globalThis as { __zzRecord?: typeof record }).__zzRecord = record;
  });

const save = (page: Page) =>
  page.evaluate(async (saves) => {
    const ref = globalThis.__folioPlayground?.getEditorRef();
    if (!ref) throw new TypeError("No editor.");
    const times: number[] = [];
    for (let index = 0; index < saves; index += 1) {
      const started = performance.now();
      const bytes = await ref.save();
      times.push(performance.now() - started);
      if (!bytes) throw new TypeError("Save returned nothing.");
    }
    return times;
  }, SAVES);

const documents = Object.keys(DOCUMENTS).filter(
  (name) => !ONLY_DOCS || ONLY_DOCS.includes(name),
);
for (const name of documents) {
  for (const mode of ["editing", "suggesting"] as const) {
    for (const engine of ORDER) {
      test(`zz flip ${name} ${mode} ${engine}`, async ({ page }) => {
        const bytes = await DOCUMENTS[name]!();
        await page.addInitScript(() => {
          const layouts: number[] = [];
          (globalThis as { __zzLayouts?: number[] }).__zzLayouts = layouts;
          globalThis.__folioLayoutInstrumentation = {
            onLayoutComplete: () => {
              layouts.push(performance.now());
            },
          };
        });
        const errors: string[] = [];
        page.on("pageerror", (error) => errors.push(error.message));
        await page.goto(`/?mode=${mode}&session=${engine}`);
        await page.waitForSelector(".layout-page");
        await page.waitForTimeout(1000);
        const opened = await open(page, bytes);
        await page.waitForTimeout(3000);
        await prepare(page);
        await page.waitForTimeout(1000);
        for (let key = 0; key < WARM_KEYS; key += 1) {
          await page.keyboard.press("x");
          await page.waitForTimeout(PAUSE_MS);
        }
        const skip = await page.evaluate(() => {
          const record = (
            globalThis as {
              __zzRecord?: {
                input: Map<number, number>;
                tasks: number[];
                frames: number[];
              };
            }
          ).__zzRecord;
          return {
            input: record?.input.size ?? 0,
            tasks: record?.tasks.length ?? 0,
            frames: record?.frames.length ?? 0,
          };
        });
        const cdp = PROFILE
          ? await page.context().newCDPSession(page)
          : undefined;
        if (cdp) {
          await cdp.send("Profiler.enable");
          await cdp.send("Profiler.setSamplingInterval", { interval: 200 });
          await cdp.send("Profiler.start");
        }
        const textBefore = await page.evaluate(
          () =>
            globalThis.__folioPlayground?.getEditorRef()?.getEditor()?.getView()
              ?.state.doc.textContent.length ?? 0,
        );
        for (let key = 0; key < MEASURED_KEYS; key += 1) {
          await page.keyboard.press("x");
          await page.waitForTimeout(PAUSE_MS);
        }
        await page.waitForTimeout(1500);
        if (cdp && PROFILE) {
          const { profile } = await cdp.send("Profiler.stop");
          mkdirSync(PROFILE, { recursive: true });
          writeFileSync(
            join(PROFILE, `${name}-${mode}-${engine}.cpuprofile`),
            JSON.stringify(profile),
          );
        }
        const textAfter = await page.evaluate(
          () =>
            globalThis.__folioPlayground?.getEditorRef()?.getEditor()?.getView()
              ?.state.doc.textContent.length ?? 0,
        );
        const measured = await page.evaluate((from) => {
          const record = (
            globalThis as {
              __zzRecord?: {
                input: Map<number, number>;
                tasks: number[];
                frames: number[];
              };
            }
          ).__zzRecord;
          return {
            input: [...(record?.input.values() ?? [])].slice(from.input),
            tasks: record?.tasks.slice(from.tasks) ?? [],
            frames: record?.frames.slice(from.frames) ?? [],
          };
        }, skip);
        const canonicalActive = await page.evaluate(
          () =>
            globalThis.__folioPlayground
              ?.getEditorRef()
              ?.getEditor()
              ?.getCanonicalDocument() != null,
        );
        const saves =
          name === "lir011" || process.env["FOLIO_ZZ_SAVE_ALL"] === "1"
            ? await save(page)
            : [];
        const line = {
          round: ROUND,
          document: name,
          mode,
          engine,
          canonicalActive,
          profiled: Boolean(PROFILE),
          blocks: opened.blocks,
          openMs: Math.round(opened.openMs),
          loadMs: Math.round(opened.loadMs),
          keys: measured.frames.length,
          typed: textAfter - textBefore,
          inputMs: median(measured.input),
          inputEntries: measured.input.length,
          taskMs: median(measured.tasks),
          frameMs: median(measured.frames),
          saveMs: median(saves),
          saves,
          errors: errors.slice(0, 3),
        };
        console.info(`ZZFLIP ${JSON.stringify(line)}`);
        if (OUT) appendFileSync(OUT, `${JSON.stringify(line)}\n`);
      });
    }
  }
}
