import { expect, test } from "bun:test";
import fc from "fast-check";

import { assertProperty, propertyTestTimeout } from "../../../../../test/property-testing";
import { ensureParaIds } from "../../docx/ensureParaIds";
import { createDocx } from "../../docx/rezip";
import { documentShape, type DocumentShape } from "../../__tests__/documentShapes";
import {
  CONFORMANCE_OPERATIONS,
  runLegacyConformanceCase,
  type ConformanceOperation,
} from "../../__tests__/editorCommandConformance";
import {
  createHarnessState,
  parseShapeDocument,
  readBack,
  resolveAllChanges,
  saveHarnessState,
  summarizeEffectiveParagraphs,
  summarizeState,
  textblocks,
} from "../../__tests__/editorHarness";
import type { SelectionPlacement } from "../../__tests__/editorHarness";

const SHAPE_IDS = [
  "single-decimal-list",
  "single-bullet-list",
  "mixed-lists",
  "outline-level-numbered",
] as const;

type ShapeId = (typeof SHAPE_IDS)[number];
type PropertyCase = { shape: ShapeId; operation: string; placement: SelectionPlacement };
const CLIPBOARD_CASES = [
  { operation: "paste:paragraphs", placement: "cross-paragraph" },
  { operation: "paste:copied-blocks", placement: "cross-paragraph" },
  { operation: "paste:single-copied-paragraph", placement: "cross-paragraph" },
] as const;

const EDIT_CASES = [
  { operation: "key:Delete", placement: "caret-end" },
  { operation: "key:Delete", placement: "cross-paragraph" },
  { operation: "command:applyStyle(Heading1)", placement: "caret-middle" },
  { operation: "command:applyStyle(Heading2)", placement: "caret-middle" },
  { operation: "command:clearStyle", placement: "caret-middle" },
] as const satisfies readonly { operation: string; placement: SelectionPlacement }[];

const PROPERTY_CASES = [
  ...SHAPE_IDS.flatMap((shape) =>
    CLIPBOARD_CASES.map(({ operation, placement }) => ({ shape, operation, placement })),
  ),
  ...SHAPE_IDS.flatMap((shape) =>
    EDIT_CASES.map(({ operation, placement }) => ({ shape, operation, placement })),
  ),
] as const satisfies readonly PropertyCase[];

type GeneratedIndentation = {
  levelLeft: readonly [number, number, number];
  levelFirstLine: readonly [number, number, number];
  authoredLeft: number | undefined;
};

const generatedIndentation = fc.record({
  levelLeft: fc.tuple(
    fc.integer({ min: 240, max: 2160 }),
    fc.integer({ min: 240, max: 2160 }),
    fc.integer({ min: 240, max: 2160 }),
  ),
  levelFirstLine: fc.tuple(
    fc.integer({ min: 0, max: 720 }),
    fc.integer({ min: 0, max: 720 }),
    fc.integer({ min: 0, max: 720 }),
  ),
  authoredLeft: fc.option(fc.constantFrom(0, 240, 720, 1440), { nil: undefined }),
});

type GeneratedShapeOptions = {
  shape: DocumentShape;
  indentation: GeneratedIndentation;
  previousChange?: "numbering";
};

const withGeneratedIndentation = async ({
  shape,
  indentation,
  previousChange,
}: GeneratedShapeOptions): Promise<DocumentShape> => {
  const build = async (): Promise<Uint8Array> => {
    const document = await parseShapeDocument(await shape.build());
    const numbering = document.package.numbering;
    if (!numbering) throw new Error(`${shape.id} has no numbering definitions`);

    for (const abstractNum of numbering.abstractNums) {
      for (const level of abstractNum.levels) {
        const index = level.ilvl % indentation.levelLeft.length;
        const indentLeft = indentation.levelLeft[index];
        const indentFirstLine = indentation.levelFirstLine[index];
        if (indentLeft === undefined || indentFirstLine === undefined) {
          throw new Error("Missing generated indentation for level " + level.ilvl);
        }
        level.pPr = {
          ...level.pPr,
          indentLeft,
          indentFirstLine,
          hangingIndent: true,
        };
      }
    }

    let listParagraphIndex = 0;
    for (const block of document.package.document.content) {
      if (block.type !== "paragraph" || block.formatting?.numPr?.kind !== "reference") continue;
      if (listParagraphIndex === 0 && indentation.authoredLeft !== undefined) {
        block.formatting = { ...block.formatting, indentLeft: indentation.authoredLeft };
      }
      listParagraphIndex += 1;
    }
    if (previousChange === "numbering") {
      const numbered = document.package.document.content.find(
        (block) => block.type === "paragraph" && block.formatting?.numPr?.kind === "reference",
      );
      const plain = document.package.document.content.find(
        (block) => block.type === "paragraph" && block.formatting?.numPr === undefined,
      );
      if (
        !numbered ||
        numbered.type !== "paragraph" ||
        numbered.formatting?.numPr?.kind !== "reference" ||
        !plain ||
        plain.type !== "paragraph"
      ) {
        throw new Error(shape.id + " needs a numbered and a plain paragraph");
      }
      plain.propertyChanges = [
        {
          type: "paragraphPropertyChange",
          info: { id: 950, author: "Property test", date: "2026-01-01T00:00:00Z" },
          previousFormatting: { numPr: numbered.formatting.numPr },
        },
      ];
    }

    const bytes = await createDocx(document);
    return (await ensureParaIds(bytes)).docx;
  };

  const id = [
    shape.id,
    ...indentation.levelLeft,
    ...indentation.levelFirstLine,
    indentation.authoredLeft ?? "inherited",
    previousChange === "numbering" ? "previous-numbering-change" : "no-change",
  ].join("-indent-");
  return { ...shape, id, build, rebuild: build };
};

const copiedParagraph: ConformanceOperation = {
  id: "paste:single-copied-paragraph",
  placements: ["cross-paragraph"],
  run: ({ view, focus }) => {
    const source = textblocks(view.state.doc).find(({ node }) => node.textContent.includes(focus));
    if (!source) return false;
    view.paste(view.state.doc.slice(source.pos + 1, source.pos + 1 + source.node.content.size));
    return true;
  },
};

const operationById = (id: string): ConformanceOperation => {
  const operation =
    id === copiedParagraph.id
      ? copiedParagraph
      : CONFORMANCE_OPERATIONS.find((candidate) => candidate.id === id);
  if (!operation) throw new Error(`Missing conformance operation ${id}`);
  return operation;
};

// Fixed list fixtures did not cross level indentation with authored overrides
// and restoration sources. Exercise each producer against the same readback oracle.
test(
  "list indentation provenance survives clipboard, Delete, and style operations",
  async () => {
    await assertProperty(
      fc.asyncProperty(generatedIndentation, async (indentation) => {
        for (const entry of PROPERTY_CASES) {
          const shape = await withGeneratedIndentation({
            shape: documentShape(entry.shape),
            indentation,
          });
          const result = await runLegacyConformanceCase({
            shape,
            operation: operationById(entry.operation),
            placement: entry.placement,
          });
          expect(result, `${entry.shape} / ${entry.operation}`).not.toBeNull();
          expect(result?.violations, `${entry.shape} / ${entry.operation}`).toEqual([]);
        }
      }),
      {
        numRuns: 8,
        id: "list indentation provenance survives clipboard, Delete, and style operations",
      },
    );
  },
  propertyTestTimeout(60_000),
);

test(
  "rejecting an imported numbering-only paragraph change restores inherited list indentation",
  async () => {
    await assertProperty(
      fc.asyncProperty(generatedIndentation, async (indentation) => {
        for (const shapeId of SHAPE_IDS) {
          const shape = await withGeneratedIndentation({
            shape: documentShape(shapeId),
            indentation,
            previousChange: "numbering",
          });
          const base = await parseShapeDocument(await shape.build());
          const before = createHarnessState(base, "suggesting");
          const rejected = resolveAllChanges(before, "reject");
          const saved = await saveHarnessState(rejected, base);
          const reopened = await readBack(saved.bytes);
          expect(summarizeState(rejected)).toEqual(reopened.summary);
          expect(summarizeEffectiveParagraphs(rejected)).toEqual(reopened.effective);
        }
      }),
      {
        numRuns: 8,
        id: "rejecting an imported numbering-only paragraph change restores inherited list indentation",
      },
    );
  },
  propertyTestTimeout(60_000),
);
