/**
 * A paragraph inserted beside another copies that paragraph's stated
 * properties as they stand when the insertion runs in the batch's execution
 * order: equal coordinates run in reverse input order, an insertion before a
 * paragraph shares its start with an edit of the paragraph's properties, and a
 * merge of the paragraph sits at its end. So an edit the order runs first is
 * in the copy, and one it runs after is not.
 */

import { describe, expect, spyOn, test } from "bun:test";

import * as documentOperations from "@stll/docx-core/ops";

import { FolioDocxReviewer } from "../ai-edits/headless";
import type {
  FolioDocumentOperation,
  FolioDocumentOperationMode,
} from "../document-operations";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { createEmptyDocument } from "../utils/createDocument";

const INSERTED = "Inserted clause.";

const open = async () => {
  const input = (
    await prepareCanonicalDocxInput(
      await createDocx(
        fromMarkdown(
          "# Terms\n\n1. Numbered clause one\n\nUnnumbered body text.\n\nClosing.",
        ),
      ),
    )
  )
    .unwrap()
    .slice().buffer;
  const reviewer = await FolioDocxReviewer.fromBuffer(input, {
    session: "canonical",
  });
  const [heading, numbered, body, closing] = reviewer.snapshot().blocks;
  if (
    heading?.text !== "Terms" ||
    numbered?.listReference === undefined ||
    body?.text !== "Unnumbered body text." ||
    closing?.text !== "Closing."
  )
    throw new Error("Missing fixture");
  return { reviewer, heading, numbered, body, closing };
};

const insertion = (
  type: "insertBeforeBlock" | "insertAfterBlock",
  blockId: string,
): FolioDocumentOperation => ({
  id: "insert",
  type,
  blockId,
  text: INSERTED,
});

const inserted = (reviewer: FolioDocxReviewer) => {
  const block = reviewer
    .snapshot()
    .blocks.find(({ text }) => text === INSERTED);
  if (block === undefined) throw new Error("The insertion is missing");
  return block;
};

const MODES = ["direct", "tracked-changes"] as const;

const apply = (
  reviewer: FolioDocxReviewer,
  operations: FolioDocumentOperation[],
  mode: FolioDocumentOperationMode,
) => {
  const result = reviewer.applyDocumentOperations(
    { version: 1, mode, atomic: true, operations },
    { undefinedReferences: "refuse" },
  );
  expect(result.skipped).toEqual([]);
  expect(result.applied).toHaveLength(operations.length);
};

describe("an inserted paragraph copies its neighbour as the batch order leaves it", () => {
  for (const mode of MODES) {
    test(`before a paragraph the batch then numbers nothing (${mode})`, async () => {
      const { reviewer, numbered } = await open();
      apply(
        reviewer,
        [
          insertion("insertBeforeBlock", numbered.id),
          {
            id: "unnumber",
            type: "setBlockParagraphProperties",
            blockId: numbered.id,
            properties: { numbering: { kind: "none" } },
          },
        ],
        mode,
      );
      expect(inserted(reviewer).listReference).toBeUndefined();
    });

    test(`before a paragraph the batch numbered nothing first, in input order (${mode})`, async () => {
      const { reviewer, numbered } = await open();
      apply(
        reviewer,
        [
          {
            id: "unnumber",
            type: "setBlockParagraphProperties",
            blockId: numbered.id,
            properties: { numbering: { kind: "none" } },
          },
          insertion("insertBeforeBlock", numbered.id),
        ],
        mode,
      );
      expect(inserted(reviewer).listReference).toEqual(numbered.listReference);
    });

    test(`after a paragraph the batch then numbers nothing (${mode})`, async () => {
      const { reviewer, numbered } = await open();
      apply(
        reviewer,
        [
          insertion("insertAfterBlock", numbered.id),
          {
            id: "unnumber",
            type: "setBlockParagraphProperties",
            blockId: numbered.id,
            properties: { numbering: { kind: "none" } },
          },
        ],
        mode,
      );
      expect(inserted(reviewer).listReference).toEqual(numbered.listReference);
    });

    test(`before a paragraph the batch then clears the style of (${mode})`, async () => {
      const { reviewer, heading } = await open();
      apply(
        reviewer,
        [
          insertion("insertBeforeBlock", heading.id),
          {
            id: "unstyle",
            type: "setBlockParagraphProperties",
            blockId: heading.id,
            properties: { styleId: null },
          },
        ],
        mode,
      );
      expect(inserted(reviewer).styleId).toBeUndefined();
    });

    test(`before a paragraph the batch then starts a list with (${mode})`, async () => {
      const { reviewer, body } = await open();
      apply(
        reviewer,
        [
          insertion("insertBeforeBlock", body.id),
          {
            id: "list",
            type: "setBlockParagraphProperties",
            blockId: body.id,
            properties: { numbering: { kind: "newList", format: "bullet" } },
          },
        ],
        mode,
      );
      const listed = reviewer
        .snapshot()
        .blocks.find(({ id }) => id === body.id)?.listReference;
      expect(listed).toBeDefined();
      expect(inserted(reviewer).listReference).toEqual(listed);
    });
  }

  // A tracked merge leaves the paragraph after it, whose properties the
  // merge sets: the paragraph inserted before the joined one copies them.
  test(`before a paragraph the batch merges with new properties (tracked-changes)`, async () => {
    const { reviewer, heading, numbered } = await open();
    const numPr = numbered.statedNumbering;
    if (numPr?.kind !== "reference") throw new Error("Missing numbering");
    apply(
      reviewer,
      [
        {
          id: "merge",
          type: "mergeBlockWithNext",
          blockId: heading.id,
          separator: " ",
          mergedParagraphProperties: { numbering: numPr },
        },
        insertion("insertBeforeBlock", heading.id),
      ],
      "tracked-changes",
    );
    expect(inserted(reviewer).listReference).toEqual(numbered.listReference);
  });
});

test.each([4, 8])(
  "admission counts the package once for %i reordered paragraph insertions",
  async (count) => {
    const { reviewer, body } = await open();
    const operations = Array.from(
      { length: count },
      (_, index) =>
        ({
          type: "insertBeforeBlock",
          id: `insert-${index}`,
          blockId: body.id,
          text: `Inserted clause ${index}.`,
        }) satisfies FolioDocumentOperation,
    );
    const census = spyOn(documentOperations, "leasedIdCensus");
    try {
      const result = reviewer.applyDocumentOperations(
        { version: 1, mode: "tracked-changes", atomic: true, operations },
        {
          undefinedReferences: "refuse",
          revisionStamp: { idSeed: 1, date: "2026-01-02T03:04:05Z" },
        },
      );
      expect(result.skipped).toEqual([]);
      expect(result.applied).toHaveLength(count);
      expect(census).toHaveBeenCalledTimes(1);
      expect(
        reviewer
          .snapshot()
          .blocks.map(({ text }) => text)
          .filter((text) => text.startsWith("Inserted clause ")),
      ).toEqual(operations.map(({ text }) => text));
    } finally {
      census.mockRestore();
    }
  },
);

test("staging reuses a verified blank replacement followed by paragraph insertions", async () => {
  const template = createEmptyDocument();
  const input = await createDocx({
    ...template,
    package: {
      ...template.package,
      document: { ...template.package.document, content: [] },
    },
  });
  const reviewer = await FolioDocxReviewer.fromBuffer(input, {
    session: "canonical",
  });
  const anchor = reviewer.snapshot().blocks.at(0);
  if (anchor === undefined || anchor.text !== "")
    throw new Error("The fixture needs its blank paragraph anchor.");
  const operations = Array.from(
    { length: 6 },
    (_, index) =>
      ({
        id: `insert-${index}`,
        type: index === 0 ? "replaceBlock" : "insertAfterBlock",
        blockId: anchor.id,
        text: `Appended clause ${index}.`,
      }) satisfies FolioDocumentOperation,
  );
  const applyOps = spyOn(documentOperations, "applyDocumentOps");
  try {
    apply(reviewer, operations, "tracked-changes");
    const insertions = applyOps.mock.calls.filter(([, ops]) =>
      ops.some(
        ({ type }) =>
          type === documentOperations.DOCUMENT_OP_TYPES.INSERT_BLOCKS,
      ),
    );
    expect(insertions).toHaveLength(operations.length - 1);
    expect(insertions.every(([, ops]) => ops.length === 1)).toBe(true);
    expect(reviewer.snapshot().blocks.map(({ text }) => text)).toEqual(
      operations.map(({ text }) => text),
    );
  } finally {
    applyOps.mockRestore();
  }
});

test.each([4, 8])(
  "compiles each of %i already ordered paragraph insertions once",
  async (count) => {
    const { reviewer, closing } = await open();
    const operations = Array.from(
      { length: count },
      (_, index) =>
        ({
          type: "insertAfterBlock",
          id: `insert-${index}`,
          blockId: closing.id,
          text: `Appended clause ${index}.`,
        }) satisfies FolioDocumentOperation,
    );
    const compiler = spyOn(documentOperations, "compileEditorIntent");
    const applyOps = spyOn(documentOperations, "applyDocumentOps");
    try {
      const result = reviewer.applyDocumentOperations(
        { version: 1, mode: "tracked-changes", atomic: true, operations },
        {
          undefinedReferences: "refuse",
          revisionStamp: { idSeed: 1, date: "2026-01-02T03:04:05Z" },
        },
      );
      expect(result.skipped).toEqual([]);
      expect(result.applied.map(({ id }) => id)).toEqual(
        operations.map(({ id }) => id),
      );
      expect(compiler).toHaveBeenCalledTimes(count);
      const appliedInsertions = applyOps.mock.calls.filter(([, ops]) =>
        ops.some(
          ({ type }) =>
            type === documentOperations.DOCUMENT_OP_TYPES.INSERT_BLOCKS,
        ),
      );
      expect(appliedInsertions).toHaveLength(count);
      expect(
        reviewer
          .snapshot()
          .blocks.slice(-count)
          .map(({ text }) => text),
      ).toEqual(operations.map(({ text }) => text));
      const primaryRevisions = result.applied.map(
        ({ revisionId }) => revisionId ?? -1,
      );
      expect(primaryRevisions.at(-1)).toBe(1);
      expect(primaryRevisions).toEqual(
        primaryRevisions.toSorted((left, right) => right - left),
      );
      const revisions = result.applied.flatMap(
        ({ revisionIds }) => revisionIds ?? [],
      );
      expect(revisions.length).toBeGreaterThanOrEqual(count);
      expect(new Set(revisions).size).toBe(revisions.length);
    } finally {
      applyOps.mockRestore();
      compiler.mockRestore();
    }
  },
);
