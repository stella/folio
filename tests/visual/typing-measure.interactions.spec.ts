/**
 * Typing measures only the blocks it changed.
 *
 * A key changes one paragraph, so a layout pass after it measures that
 * paragraph and reuses the measure of every other paragraph. The layout
 * instrumentation reports each block measured and each measure reused; this
 * spec types in long documents (a numbered list, a contract with lists and
 * tables, a document of tables) in both hosts and both engines, with track
 * changes off and on, and bounds the paragraphs each layout pass measures.
 * Every key waits for the layout pass it causes before the next one, so each
 * key lays out on its own whatever the machine's speed.
 *
 * A table the edit did not reach keeps its measure too, whether the pass
 * hands it out again or moves it. In the default engine a settled edit builds
 * the document again, with new styles, so the next pass measures every table
 * once: a table is measured at most once per such build.
 */

import { createDocx } from "../../packages/core/src/docx/rezip";
import type { LayoutInstrumentation } from "../../packages/core/src/layout-engine/layoutInstrumentation";
import { fromMarkdown } from "../../packages/core/src/markdown/fromMarkdown";
import type { DocxEditorRef } from "../../packages/react/src/components/DocxEditor.props";
import { PLAYGROUND_HOSTS } from "../parity/playgroundHosts";
import { expect, type Page, test } from "../support/unclassifiedProbe";

type MeasureCounts = {
  layouts: number;
  /** Keys pressed so far. */
  keys: number;
  /** The keys pressed when the layout pass under way started. */
  keysAtStart: number;
  /** The keys laid out: those pressed before the last edit's pass started. */
  keysLaidOut: number;
  measured: number;
  measuredTables: number;
  reused: number;
  conversions: number;
  /** The paragraphs each completed layout pass measured, in order. */
  passes: number[];
  /** The paragraphs the layout pass under way has measured so far. */
  current: number | undefined;
};

declare global {
  var __folioPlayground:
    { getEditorRef: () => DocxEditorRef | null } | undefined;
  var __folioLayoutInstrumentation: LayoutInstrumentation | undefined;
  var __folioMeasureCounts: MeasureCounts | undefined;
}

const KEYS = 8;
// The edited paragraph, plus a paragraph on either side whose measure reads it.
const MEASURED_PARAGRAPHS_PER_LAYOUT = 3;

const WORDS =
  "the supplier shall deliver goods under this agreement subject to its terms and the warranty".split(
    " ",
  );
const sentence = (seed: number, length: number) =>
  Array.from(
    { length },
    (_, index) => WORDS[(seed * 7 + index * 3) % WORDS.length] ?? "the",
  ).join(" ");

const DOCUMENTS = {
  "numbered list": () =>
    Array.from({ length: 300 }, (_, index) => `1. Item ${index + 1}`).join(
      "\n",
    ),
  contract: () => {
    const lines: string[] = [];
    for (let article = 1; article <= 12; article += 1) {
      lines.push(`## ${article}. ${sentence(article, 3)}`, "");
      for (let clause = 1; clause <= 6; clause += 1) {
        lines.push(`${article}.${clause} ${sentence(clause, 30)}.`, "");
        if (clause % 3 === 0) {
          for (let item = 1; item <= 3; item += 1)
            lines.push(`${item}. ${sentence(item, 12)};`);
          lines.push("");
        }
      }
      if (article % 4 === 0) {
        lines.push("| Item | Fee | Term |", "|---|---|---|");
        for (let row = 0; row < 6; row += 1)
          lines.push(
            `| ${sentence(row, 3)} | ${row * 100} | ${sentence(row, 2)} |`,
          );
        lines.push("");
      }
    }
    return lines.join("\n");
  },
  tables: () => {
    const lines: string[] = [];
    for (let table = 1; table <= 20; table += 1) {
      lines.push(`Table ${table} lists ${sentence(table, 10)}.`, "");
      lines.push("| Name | Amount | Note |", "|---|---|---|");
      for (let row = 0; row < 5; row += 1)
        lines.push(`| ${sentence(row, 2)} | ${row} | ${sentence(row, 4)} |`);
      lines.push("");
    }
    return lines.join("\n");
  },
} as const;

const HOSTS = [
  { name: "react", origin: PLAYGROUND_HOSTS.react },
  { name: "vue", origin: PLAYGROUND_HOSTS.vue },
] as const;
const ENGINES = ["default", "canonical"] as const;
const MODES = ["editing", "suggesting"] as const;

test.use({ viewport: { width: 1280, height: 900 } });

/** Put the caret in the middle of a paragraph halfway through the document. */
const placeCaret = (page: Page) =>
  page.evaluate(() => {
    const ref = globalThis.__folioPlayground?.getEditorRef();
    const editor = ref?.getEditor();
    const view = editor?.getView();
    if (!ref || !editor || !view) throw new TypeError("No editor view.");
    const paragraphs: { position: number; paraId: unknown }[] = [];
    view.state.doc.descendants((node, position) => {
      if (node.type.name === "paragraph" && node.textContent.length > 4)
        paragraphs.push({ position, paraId: node.attrs["paraId"] });
      return node.type.name !== "paragraph";
    });
    const middle = paragraphs[Math.floor(paragraphs.length / 2)];
    if (!middle || typeof middle.paraId !== "string")
      throw new TypeError("No paragraph to type in.");
    if (!ref.scrollToParaId(middle.paraId))
      throw new TypeError("No paragraph to scroll to.");
    editor.setSelection(middle.position + 3, middle.position + 3);
    editor.focus();
  });

for (const [name, markdown] of Object.entries(DOCUMENTS)) {
  for (const host of HOSTS) {
    for (const engine of ENGINES) {
      for (const mode of MODES) {
        test(`typing in a ${name} measures only the edited block [${host.name}, ${engine}, ${mode}]`, async ({
          page,
        }) => {
          test.setTimeout(120_000);
          const bytes = Buffer.from(await createDocx(fromMarkdown(markdown())));
          await page.route("**/fixtures/typing-measure.docx", (route) =>
            route.fulfill({ body: bytes }),
          );
          await page.addInitScript(() => {
            const counts: MeasureCounts = {
              layouts: 0,
              keys: 0,
              keysAtStart: 0,
              keysLaidOut: 0,
              measured: 0,
              measuredTables: 0,
              reused: 0,
              conversions: 0,
              passes: [],
              current: undefined,
            };
            globalThis.__folioMeasureCounts = counts;
            // A key's edit lands while its keydown is handled, before any
            // layout pass the key schedules can start.
            globalThis.addEventListener(
              "keydown",
              () => {
                counts.keys += 1;
              },
              true,
            );
            globalThis.__folioLayoutInstrumentation = {
              onLayoutStart: () => {
                counts.current = 0;
                counts.keysAtStart = counts.keys;
              },
              onLayoutComplete: ({ reason }) => {
                counts.layouts += 1;
                if (reason === "transaction")
                  counts.keysLaidOut = counts.keysAtStart;
                counts.passes.push(counts.current ?? 0);
                counts.current = undefined;
              },
              onLayoutError: () => {
                counts.current = undefined;
              },
              onMeasureBlock: ({ blockKind }) => {
                if (blockKind === "table") counts.measuredTables += 1;
                else {
                  counts.measured += 1;
                  if (counts.current !== undefined) counts.current += 1;
                }
              },
              onReuseMeasure: () => {
                counts.reused += 1;
              },
              onDocumentBuilt: ({ kind }) => {
                if (kind === "conversion") counts.conversions += 1;
              },
            };
          });
          const query = new URLSearchParams({
            file: "typing-measure.docx",
            mode,
          });
          if (engine === "canonical") query.set("session", "canonical");
          // The spec needs the editor and its document, not every resource
          // the page loads: it waits for those below.
          await page.goto(`${host.origin}/?${query.toString()}`, {
            waitUntil: "domcontentloaded",
          });
          await page.waitForFunction(
            () => {
              const ref = globalThis.__folioPlayground?.getEditorRef();
              ref?.ensureEditorView({ focus: false });
              return (
                (ref?.getEditor()?.getView()?.state.doc.childCount ?? 0) > 20
              );
            },
            undefined,
            { timeout: 60_000 },
          );
          await page.evaluate(() => document.fonts.ready);
          await placeCaret(page);
          let pressed = 0;
          /** Press a key, then wait for the layout pass the key causes. */
          const pressAndLayOut = async () => {
            await page.keyboard.press("x");
            pressed += 1;
            await page.waitForFunction(
              (key) =>
                (globalThis.__folioMeasureCounts?.keysLaidOut ?? 0) >= key,
              pressed,
              { timeout: 10_000 },
            );
          };
          const snapshot = () =>
            page.evaluate(() => {
              const counts = globalThis.__folioMeasureCounts;
              if (!counts) throw new TypeError("No measure counts.");
              return { ...counts, passes: [...counts.passes] };
            });
          // The first key may convert the paragraph into a tracked one.
          await pressAndLayOut();
          const before = await snapshot();
          for (let key = 0; key < KEYS; key += 1) await pressAndLayOut();
          const after = await snapshot();
          const layouts = after.layouts - before.layouts;
          const keysLaidOut = after.keysLaidOut - before.keysLaidOut;
          const passes = after.passes.slice(before.passes.length);
          const measured = after.measured - before.measured;
          const reused = after.reused - before.reused;
          const measuredTables = after.measuredTables - before.measuredTables;
          const conversions = after.conversions - before.conversions;
          const tables = await page.evaluate(() => {
            const view = globalThis.__folioPlayground
              ?.getEditorRef()
              ?.getEditor()
              ?.getView();
            let count = 0;
            view?.state.doc.descendants((node) => {
              if (node.type.name === "table") count += 1;
              return node.type.name !== "table";
            });
            return count;
          });
          console.info(
            JSON.stringify({
              name,
              host: host.name,
              engine,
              mode,
              layouts,
              keysLaidOut,
              measured,
              mostMeasuredInAPass: Math.max(0, ...passes),
              measuredTables,
              tables,
              conversions,
              reused,
            }),
          );
          expect(layouts).toBeGreaterThanOrEqual(KEYS);
          // Every other block's measure is reused, not taken again.
          expect(reused).toBeGreaterThan(layouts * 10);
          // Every pass, not only their average, stays within the bound.
          for (const pass of passes)
            expect(pass).toBeLessThanOrEqual(MEASURED_PARAGRAPHS_PER_LAYOUT);
          // A canonical session builds no document between keys, so no
          // table is measured again. The default engine measures each table
          // at most once per build: those in the window, and the one that may
          // land just before it.
          expect(measuredTables).toBeLessThanOrEqual(
            engine === "canonical" ? 0 : tables * (conversions + 1),
          );
        });
      }
    }
  }
}
