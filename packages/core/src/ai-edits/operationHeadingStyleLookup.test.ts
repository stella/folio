import { expect, test } from "bun:test";

import type { FolioDocumentOperation } from "../document-operations";
import { ensureParaIds } from "../docx/ensureParaIds";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import type { StyleDefinitions } from "../types/document";
import { FolioDocxReviewer } from "./headless";
import {
  createOperationHeadingStyles,
  withOperationHeadingStyles,
} from "./undefinedReferences";

const BLOCK_ID = "00000001";

const insert = (styleId?: string): FolioDocumentOperation => ({
  id: "insert",
  type: "insertAfterBlock",
  blockId: BLOCK_ID,
  text: "Inserted clause.",
  ...(styleId === undefined ? {} : { styleId }),
});

test("scans style IDs once per definitions object and reuses heading additions", () => {
  let reads = 0;
  const size = 256;
  const operations = 128;
  const styles: StyleDefinitions = {
    styles: Array.from({ length: size }, (_, index) => ({
      get styleId() {
        reads += 1;
        return index === 0 ? "Normal" : `Custom${index}`;
      },
      type: "paragraph",
    })),
  };
  const headings = createOperationHeadingStyles();
  for (let index = 0; index < operations; index += 1)
    expect(headings(styles, insert())).toBe(styles);
  expect(reads).toBe(0);

  const first = headings(styles, insert("Heading9"));
  expect(first?.styles.at(-1)).toMatchObject({
    styleId: "Heading9",
    basedOn: "Normal",
    next: "Normal",
  });
  for (let index = 0; index < operations; index += 1) {
    expect(headings(styles, insert("Heading9"))).toBe(first);
    expect(headings(styles, insert("Custom255"))).toBe(styles);
  }
  expect(reads).toBe(size);

  for (let index = 0; index < operations; index += 1)
    expect(headings(first, insert("Heading9"))).toBe(first);
  expect(reads).toBe(2 * size);
});

test("matches the uncached heading addition and keeps an occupied character-style ID", () => {
  const styles: StyleDefinitions = {
    styles: [
      { styleId: "Normal", type: "character" },
      { styleId: "Heading8", type: "character" },
    ],
  };
  const split = {
    id: "split",
    type: "splitBlock",
    blockId: BLOCK_ID,
    offset: 1,
    firstParagraphProperties: { styleId: "Heading8" },
    secondParagraphProperties: { styleId: "Heading2" },
  } as const satisfies FolioDocumentOperation;
  const headings = createOperationHeadingStyles();
  const augmented = headings(styles, split);
  expect(augmented).toStrictEqual(withOperationHeadingStyles(styles, [split]));
  expect(augmented?.styles).toHaveLength(3);
  expect(augmented?.styles.at(-1)?.styleId).toBe("Heading2");
  expect(augmented?.styles.at(-1)?.basedOn).toBeUndefined();
  expect(augmented?.styles.at(-1)?.next).toBeUndefined();
  expect(headings(styles, insert("Heading8"))).toBe(styles);
});

test("canonical admission and compilation keep explicit and inherited heading styles", async () => {
  const { docx } = await ensureParaIds(
    await createDocx(fromMarkdown("# Terms")),
  );
  const reviewer = await FolioDocxReviewer.fromBuffer(docx, {
    session: "canonical",
  });
  const anchor = reviewer.snapshot().blocks.at(0);
  if (anchor === undefined)
    throw new Error("The heading fixture needs an anchor.");
  const result = reviewer.applyDocumentOperations(
    {
      version: 1,
      mode: "direct",
      operations: [
        {
          type: "insertBeforeBlock",
          id: "explicit",
          blockId: anchor.id,
          text: "Explicit",
          styleId: "Heading9",
        },
        {
          type: "insertBeforeBlock",
          id: "inherited",
          blockId: anchor.id,
          text: "Inherited",
        },
      ],
    },
    { undefinedReferences: "refuse" },
  );
  expect(result.skipped).toEqual([]);
  expect(result.applied).toHaveLength(2);
  const styles = new Map(
    reviewer
      .snapshot()
      .blocks.flatMap((block) =>
        block.kind === "diagnostic" ? [] : [[block.text, block.styleId]],
      ),
  );
  expect(styles.get("Explicit")).toBe("Heading9");
  expect(styles.get("Inherited")).toBe("Heading1");
  expect(styles.get("Terms")).toBe("Heading1");
});
