/**
 * Typing measures only the paragraphs it changed.
 *
 * A key changes one paragraph, so a layout pass after it measures that
 * paragraph and reuses the measure of every other paragraph. The layout
 * instrumentation reports each block measured and each measure reused; this
 * spec types in long documents (a numbered list, a contract with lists and
 * tables, a document of tables) in both hosts and both engines, with track
 * changes off and on, and bounds the paragraphs measured per layout pass.
 *
 * Tables are converted again by every pass, so each one is measured again
 * (its cell paragraphs come from the cache). Their count is reported, not
 * bounded: reusing a table's measure needs tables that keep their identity
 * across passes, as paragraphs do.
 */

import { expect, type Page, test } from "@playwright/test";

import { createDocx } from "../../packages/core/src/docx/rezip";
import type { LayoutInstrumentation } from "../../packages/core/src/layout-engine/layoutInstrumentation";
import { fromMarkdown } from "../../packages/core/src/markdown/fromMarkdown";
import type { DocxEditorRef } from "../../packages/react/src/components/DocxEditor.props";
import { PLAYGROUND_HOSTS } from "../parity/playgroundHosts";

type MeasureCounts = {
  layouts: number;
  measured: number;
  measuredTables: number;
  reused: number;
};

declare global {
  var __folioPlayground:
    { getEditorRef: () => DocxEditorRef | null } | undefined;
  var __folioLayoutInstrumentation: LayoutInstrumentation | undefined;
  var __folioMeasureCounts: MeasureCounts | undefined;
}

const KEYS = 8;
// Long enough for each key to lay out on its own.
const PAUSE_MS = 200;
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
              measured: 0,
              measuredTables: 0,
              reused: 0,
            };
            globalThis.__folioMeasureCounts = counts;
            globalThis.__folioLayoutInstrumentation = {
              onLayoutComplete: () => {
                counts.layouts += 1;
              },
              onMeasureBlock: ({ blockKind }) => {
                if (blockKind === "table") counts.measuredTables += 1;
                else counts.measured += 1;
              },
              onReuseMeasure: () => {
                counts.reused += 1;
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
          // The first key may convert the paragraph into a tracked one.
          await page.keyboard.press("x");
          await page.waitForTimeout(1_000);
          const before = await page.evaluate(() => ({
            ...globalThis.__folioMeasureCounts,
          }));
          for (let key = 0; key < KEYS; key += 1) {
            await page.keyboard.press("x");
            await page.waitForTimeout(PAUSE_MS);
          }
          const after = await page.evaluate(() => ({
            ...globalThis.__folioMeasureCounts,
          }));
          const layouts = (after.layouts ?? 0) - (before.layouts ?? 0);
          const measured = (after.measured ?? 0) - (before.measured ?? 0);
          const reused = (after.reused ?? 0) - (before.reused ?? 0);
          const measuredTables =
            (after.measuredTables ?? 0) - (before.measuredTables ?? 0);
          console.info(
            JSON.stringify({
              name,
              host: host.name,
              engine,
              mode,
              layouts,
              measured,
              measuredTables,
              reused,
            }),
          );
          expect(layouts).toBeGreaterThanOrEqual(KEYS);
          // Every other block's measure is reused, not taken again.
          expect(reused).toBeGreaterThan(layouts * 10);
          expect(measured).toBeLessThanOrEqual(
            layouts * MEASURED_PARAGRAPHS_PER_LAYOUT,
          );
        });
      }
    }
  }
}
