/**
 * A table's reused measure equals measuring it again.
 *
 * Random runs of typing, deleting and splitting paragraphs before, inside and
 * after tables are converted with kept blocks after every edit, so tables an
 * edit did not reach come back as the same blocks or moved. Every table's
 * measure, as measuring hands it out, must equal measuring that table afresh,
 * and some must have been reused.
 */
import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { TextSelection } from "prosemirror-state";

import {
  propertyConfig,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import { DOCUMENT_SHAPES } from "../../__tests__/documentShapes";
import {
  createHarnessState,
  HeadlessEditorView,
  parseShapeDocument,
} from "../../__tests__/editorHarness";
import { createDocx } from "../../docx/rezip";
import type { LayoutInstrumentation } from "../../layout-engine/layoutInstrumentation";
import {
  fixedCharWidth,
  withFakeTextMeasure,
} from "../../layout-engine/measure/__tests__/fakeTextMeasure";
import {
  measureBlocks,
  measureTableBlock,
} from "../../layout-engine/measure/measureBlocks";
import type { FlowBlock, TableBlock } from "../../layout-engine/types";
import { fromMarkdown } from "../../markdown/fromMarkdown";
import { ParagraphBlockCache } from "./paragraphBlockCache";
import { toFlowBlocks, WORD_UNSPECIFIED_FONT_SIZE } from "./toFlowBlocks";

setDefaultTimeout(propertyTestTimeout(120_000));

const WIDTHS = [480, 360] as const;

const editArbitrary = fc.record({
  kind: fc.constantFrom("type", "delete", "enter", "width", "cell"),
  at: fc.double({ min: 0, max: 1, noNaN: true }),
  text: fc.constantFrom("x", "new words"),
});
type Edit = fc.RecordValue<typeof editArbitrary>;

/** Positions inside a table cell's text, where a cell edit types. */
const cellPositions = (view: HeadlessEditorView): number[] => {
  const positions: number[] = [];
  view.state.doc.descendants((node, pos, parent) => {
    if (
      node.isTextblock &&
      (parent?.type.name === "tableCell" || parent?.type.name === "tableHeader")
    )
      positions.push(pos + 1);
    return true;
  });
  return positions;
};

const applyEdit = (view: HeadlessEditorView, edit: Edit): void => {
  if (edit.kind === "width") return;
  const { doc } = view.state;
  const cells = edit.kind === "cell" ? cellPositions(view) : [];
  const at =
    edit.kind === "cell" && cells.length > 0
      ? (cells[Math.floor(edit.at * cells.length)] ?? 0)
      : Math.floor(edit.at * doc.content.size);
  view.dispatch(
    view.state.tr.setSelection(TextSelection.near(doc.resolve(at))),
  );
  switch (edit.kind) {
    case "type":
    case "cell":
      view.typeText(edit.text);
      return;
    case "delete":
      view.pressKey("Backspace");
      return;
    case "enter":
      view.pressKey("Enter");
      return;
    case "width":
      return;
    default: {
      const unhandled: never = edit.kind;
      throw new TypeError(`Unhandled edit ${String(unhandled)}`);
    }
  }
};

const tablesIn = (blocks: readonly FlowBlock[]): TableBlock[] =>
  blocks.filter((block): block is TableBlock => block.kind === "table");

test("a reused table measure equals measuring the table again", async () => {
  const shapes = DOCUMENT_SHAPES.filter(({ id }) =>
    ["tables", "comments", "tracked-changes"].includes(id),
  );
  // Tables without a grid share the text column's width out between them.
  const gridless = [
    "Fees under this agreement.",
    "",
    "| Item | Fee | Term |",
    "|---|---|---|",
    "| delivery of goods under the agreement | 100 | thirty days |",
    "| support | 200 | one year |",
    "",
    "Terms that follow the table.",
  ].join("\n");
  const documents = await Promise.all([
    ...shapes.map(async ({ build }) => parseShapeDocument(await build())),
    (async () =>
      parseShapeDocument(
        new Uint8Array(await createDocx(fromMarkdown(gridless))),
      ))(),
  ]);
  let reused = 0;
  const previous = globalThis.__folioLayoutInstrumentation;
  globalThis.__folioLayoutInstrumentation = {
    onReuseMeasure: ({ blockKind }) => {
      if (blockKind === "table") reused += 1;
    },
  } satisfies LayoutInstrumentation;
  try {
    withFakeTextMeasure(
      () => {
        fc.assert(
          fc.property(
            fc.integer({ min: 0, max: documents.length - 1 }),
            fc.array(editArbitrary, { minLength: 1, maxLength: 8 }),
            (which, edits) => {
              const document = documents[which]!;
              const view = new HeadlessEditorView(
                createHarnessState(document, "editing"),
              );
              const cache = new ParagraphBlockCache();
              const convert = () =>
                toFlowBlocks(view.state.doc, {
                  pageContentHeight: 900,
                  styles: document.package.styles,
                  theme: document.package.theme ?? null,
                  defaultSize: WORD_UNSPECIFIED_FONT_SIZE,
                  paragraphBlocks: cache,
                });
              let width: number = WIDTHS[0];
              const check = () => {
                for (const table of tablesIn(convert())) {
                  // A layout pass's path for a table, which may reuse.
                  const handedOut = measureBlocks([table], width).at(0);
                  expect(handedOut).toEqual(measureTableBlock(table, width));
                }
              };
              check();
              for (const edit of edits) {
                // The text column changes width, as a page setup change does.
                if (edit.kind === "width")
                  width = width === WIDTHS[0] ? WIDTHS[1] : WIDTHS[0];
                applyEdit(view, edit);
                check();
              }
            },
          ),
          propertyConfig({ numRuns: 60 }),
        );
      },
      { charWidth: fixedCharWidth(5) },
    );
  } finally {
    globalThis.__folioLayoutInstrumentation = previous;
  }
  // Tables an edit did not reach took their measure again.
  expect(reused).toBeGreaterThan(0);
});
