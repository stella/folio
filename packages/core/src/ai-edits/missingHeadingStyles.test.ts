import { expect, test } from "bun:test";

import type { FolioDocumentOperation } from "../document-operations";
import { ensureParaIds } from "../docx/ensureParaIds";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { FolioDocxReviewer } from "./headless";

const fixture = async () => {
  const document = fromMarkdown("First clause.\n\nSecond clause.");
  document.package.styles = {
    styles: [
      { styleId: "Normal", type: "paragraph", default: true },
      { styleId: "Heading8", type: "character", rPr: { italic: true } },
    ],
  };
  return (await ensureParaIds(await createDocx(document))).docx;
};

const findStyle = (reviewer: FolioDocxReviewer, id: string) =>
  reviewer
    .toDocument()
    .package.styles?.styles.find(({ styleId }) => styleId === id);

const styleOperation = (id: string, blockId: string, styleId: string) =>
  ({
    id,
    type: "setBlockParagraphProperties",
    blockId,
    properties: { styleId },
  }) as const satisfies FolioDocumentOperation;

for (const session of ["legacy", "canonical"] as const) {
  test(`materializes one exact heading definition on apply in ${session}`, async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await fixture(), {
      session,
    });
    const [first, second] = reviewer.snapshot().blocks;
    if (!first || !second) throw new Error("Expected two fixture paragraphs");
    const batch = {
      version: 1,
      mode: "direct",
      operations: [
        styleOperation("one", first.id, "Heading9"),
        styleOperation("two", second.id, "Heading9"),
      ],
    } as const;
    const options = { undefinedReferences: "refuse" } as const;
    expect(
      reviewer.applyDocumentOperations({ ...batch, dryRun: true }, options)
        .applied,
    ).toHaveLength(2);
    expect(findStyle(reviewer, "Heading9")).toBeUndefined();

    const result = reviewer.applyDocumentOperations(batch, options);
    expect(result.skipped).toEqual([]);
    expect(result.applied).toHaveLength(2);
    const reopened = await FolioDocxReviewer.fromBuffer(
      await reviewer.toBuffer(),
      { session },
    );
    expect(findStyle(reopened, "Heading9")).toMatchObject({
      type: "paragraph",
      pPr: { outlineLevel: { kind: "heading", level: 8 } },
    });
    expect(
      reopened
        .toDocument()
        .package.styles?.styles.filter(({ styleId }) => styleId === "Heading9"),
    ).toHaveLength(1);
    expect(
      reopened
        .snapshot()
        .blocks.map((block) =>
          block.kind === "diagnostic" ? undefined : block.styleId,
        ),
    ).toEqual(["Heading9", "Heading9"]);
    if (result.undoHandle === null) throw new Error("Expected an undo handle");
    expect(reviewer.undoDocumentOperations(result.undoHandle).status).toBe(
      "undone",
    );
    expect(findStyle(reviewer, "Heading9")).toBeUndefined();
  });

  test(`leaves the stylesheet unchanged on refused edits in ${session}`, async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await fixture(), {
      session,
    });
    const [first, second] = reviewer.snapshot().blocks;
    if (!first || !second) throw new Error("Expected two fixture paragraphs");
    const result = reviewer.applyDocumentOperations(
      {
        version: 1,
        mode: "direct",
        atomic: true,
        operations: [
          styleOperation("heading", first.id, "Heading1"),
          styleOperation("wrongType", second.id, "Heading8"),
        ],
      },
      { undefinedReferences: "refuse" },
    );
    expect(result.applied).toEqual([]);
    expect(result.skipped).toContainEqual(
      expect.objectContaining({ id: "wrongType", reason: "missingStyle" }),
    );
    expect(findStyle(reviewer, "Heading1")).toBeUndefined();
    expect(findStyle(reviewer, "Heading8")?.type).toBe("character");

    const bestEffort = reviewer.applyDocumentOperations(
      {
        version: 1,
        mode: "direct",
        operations: [
          styleOperation("missingTarget", "10000000", "Heading7"),
          styleOperation("custom", first.id, "heading1"),
          styleOperation("heading", second.id, "Heading1"),
        ],
      },
      { undefinedReferences: "refuse" },
    );
    expect(bestEffort.applied.map(({ id }) => id)).toEqual(["heading"]);
    expect(findStyle(reviewer, "Heading7")).toBeUndefined();
    expect(findStyle(reviewer, "Heading1")?.type).toBe("paragraph");
  });

  test(`stages headings with insertions sharing a gap in ${session}`, async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await fixture(), {
      session,
    });
    const anchor = reviewer.snapshot().blocks.at(0);
    if (anchor === undefined) throw new Error("Expected an insertion anchor");
    const result = reviewer.applyDocumentOperations(
      {
        version: 1,
        mode: "direct",
        operations: ["Heading1", "Heading7", "Heading1"].map(
          (styleId, index) =>
            ({
              id: `insert-${String(index)}`,
              type: "insertAfterBlock",
              blockId: anchor.id,
              styleId,
              text: `Inserted clause ${String(index + 1)}.`,
            }) satisfies FolioDocumentOperation,
        ),
      },
      { undefinedReferences: "refuse" },
    );
    expect(result.skipped).toEqual([]);
    expect(result.applied).toHaveLength(3);
    expect(
      reviewer
        .toDocument()
        .package.styles?.styles.filter(
          ({ styleId }) => styleId === "Heading1" || styleId === "Heading7",
        )
        .map(({ styleId }) => styleId)
        .sort(),
    ).toEqual(["Heading1", "Heading7"]);
    const undo = result.undoHandle;
    if (undo === null) throw new Error("Expected an undo handle");
    expect(reviewer.undoDocumentOperations(undo).status).toBe("undone");
    expect(findStyle(reviewer, "Heading1")).toBeUndefined();
    expect(findStyle(reviewer, "Heading7")).toBeUndefined();
  });
}
