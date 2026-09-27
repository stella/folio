/**
 * Paragraph styles the document does not define.
 *
 * A `styleId` naming no `w:style`, or a table or character style, used to
 * apply and report success: the paragraph reopened with that `w:pStyle` and
 * kept its body formatting, since only a paragraph-type definition confers
 * any. It is now skipped before anything is applied, like a numbering
 * instance the document does not define.
 */

import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { ensureParaIds } from "../docx/ensureParaIds";
import { createDocx } from "../docx/rezip";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  FOLIO_DOCUMENT_OPERATION_MODES,
  isFolioDocumentOperationModeSupported,
  type FolioDocumentOperation,
  type FolioDocumentOperationMode,
} from "../document-operations";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { FolioDocxReviewer } from "./headless";

const HEADING = "Agreement";
const ANCHOR = "First clause.";
const TAIL = "Tail.";
const MARKDOWN = `# ${HEADING}\n\n${ANCHOR}\n\n${TAIL}`;

const build = async (): Promise<Uint8Array> => {
  const model = fromMarkdown(MARKDOWN);
  // A character style next to the table styles the markdown package defines.
  model.package.styles?.styles.push({
    styleId: "Strong",
    type: "character",
    name: "Strong",
    rPr: { bold: true },
  });
  return (await ensureParaIds(await createDocx(model))).docx;
};

type Ids = { heading: string; anchor: string; tail: string };

const openReviewer = async () => {
  const reviewer = await FolioDocxReviewer.fromBuffer(await build(), { author: "AI" });
  const blocks = reviewer.getContent();
  const find = (text: string) => {
    const id = blocks.find((block) => block.text === text)?.id;
    if (id === undefined) {
      throw new Error(`fixture must expose "${text}"`);
    }
    return id;
  };
  return { reviewer, ids: { heading: find(HEADING), anchor: find(ANCHOR), tail: find(TAIL) } };
};

type StyleOperationFactory = (ids: Ids, styleId: string | null) => FolioDocumentOperation;

/** Every operation that carries a paragraph style, by where it carries it. */
const STYLE_OPERATIONS: Record<string, StyleOperationFactory> = {
  insertAfterBlock: ({ anchor }, styleId) => ({
    id: "op",
    type: "insertAfterBlock",
    blockId: anchor,
    text: "Inserted clause.",
    styleId,
  }),
  insertBeforeBlock: ({ anchor }, styleId) => ({
    id: "op",
    type: "insertBeforeBlock",
    blockId: anchor,
    text: "Inserted clause.",
    styleId,
  }),
  replaceBlock: ({ anchor }, styleId) => ({
    id: "op",
    type: "replaceBlock",
    blockId: anchor,
    text: "Rewritten clause.",
    styleId,
  }),
  setBlockParagraphProperties: ({ anchor }, styleId) => ({
    id: "op",
    type: "setBlockParagraphProperties",
    blockId: anchor,
    properties: { styleId },
  }),
  "splitBlock (first paragraph)": ({ anchor }, styleId) => ({
    id: "op",
    type: "splitBlock",
    blockId: anchor,
    offset: 6,
    firstParagraphProperties: { styleId },
  }),
  "splitBlock (second paragraph)": ({ anchor }, styleId) => ({
    id: "op",
    type: "splitBlock",
    blockId: anchor,
    offset: 6,
    secondParagraphProperties: { styleId },
  }),
  mergeBlockWithNext: ({ anchor }, styleId) => ({
    id: "op",
    type: "mergeBlockWithNext",
    blockId: anchor,
    mergedParagraphProperties: { styleId },
  }),
};

const UNDEFINED_STYLES = {
  "an undefined id": { styleId: "NoSuchStyle", names: "names no style in this document" },
  "a table style": { styleId: "TableGrid", names: "names a table style, not a paragraph style" },
  "a character style": {
    styleId: "Strong",
    names: "names a character style, not a paragraph style",
  },
} as const;

const supportedModes = (factory: StyleOperationFactory) =>
  FOLIO_DOCUMENT_OPERATION_MODES.filter((mode) =>
    isFolioDocumentOperationModeSupported(
      factory({ heading: "", anchor: "", tail: "" }, null).type,
      mode,
    ),
  );

const cases = Object.entries(UNDEFINED_STYLES).flatMap(([styleName, { styleId, names }]) =>
  Object.entries(STYLE_OPERATIONS).flatMap(([operationName, factory]) =>
    supportedModes(factory).map((mode) => ({
      styleName,
      styleId,
      names,
      operationName,
      factory,
      mode,
    })),
  ),
);

const applyOne = (
  reviewer: FolioDocxReviewer,
  operation: FolioDocumentOperation,
  mode: FolioDocumentOperationMode,
) =>
  reviewer.applyDocumentOperations({
    version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
    mode,
    operations: [operation],
  });

const texts = (reviewer: FolioDocxReviewer) => reviewer.getContent().map((block) => block.text);

const documentXml = async (docx: Uint8Array): Promise<string> =>
  (await (await JSZip.loadAsync(docx)).file("word/document.xml")?.async("text")) ?? "";

describe("a paragraph style the document does not define is skipped before anything is applied", () => {
  test.each(cases)(
    "$operationName with $styleName, $mode mode",
    async ({ styleId, names, factory, mode }) => {
      const { reviewer, ids } = await openReviewer();
      const before = texts(reviewer);

      const result = applyOne(reviewer, factory(ids, styleId), mode);

      expect(result.applied).toEqual([]);
      expect(result.receipts).toEqual([]);
      expect(result.skipped).toEqual([
        { id: "op", reason: "missingStyle", message: expect.stringContaining(names) },
      ]);
      expect(result.issues).toEqual([
        {
          operationId: "op",
          operationIndex: 0,
          path: "$.operations[0]",
          code: "missingStyle",
          retryable: true,
          recovery: "refreshDocument",
          message: expect.stringContaining(`"${styleId}"`),
        },
      ]);
      expect(texts(reviewer)).toEqual(before);
      const saved = await reviewer.toBuffer();
      expect(await documentXml(saved)).not.toContain(`w:val="${styleId}"`);
    },
  );

  test("the skip names the field and the paragraph styles the document offers", async () => {
    const { reviewer, ids } = await openReviewer();
    const result = applyOne(
      reviewer,
      STYLE_OPERATIONS["splitBlock (second paragraph)"]!(ids, "NoSuchStyle"),
      "direct",
    );
    const message = result.skipped[0]?.message ?? "";
    expect(message).toStartWith('secondParagraphProperties.styleId "NoSuchStyle"');
    expect(message).toContain("Heading2");
    expect(message).not.toContain("TableGrid");
  });

  test("a batch keeps its other operations", async () => {
    const { reviewer, ids } = await openReviewer();
    const result = reviewer.applyDocumentOperations({
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "tracked-changes",
      operations: [
        STYLE_OPERATIONS["setBlockParagraphProperties"]!(ids, "NoSuchStyle"),
        { id: "tail", type: "replaceBlock", blockId: ids.tail, text: "New tail." },
      ],
    });
    expect(result.applied.map(({ id }) => id)).toEqual(["tail"]);
    expect(result.skipped.map(({ reason }) => reason)).toEqual(["missingStyle"]);
  });

  test("an atomic batch applies none of its operations", async () => {
    const { reviewer, ids } = await openReviewer();
    const before = texts(reviewer);
    const result = reviewer.applyDocumentOperations({
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "tracked-changes",
      atomic: true,
      operations: [
        { id: "tail", type: "replaceBlock", blockId: ids.tail, text: "New tail." },
        STYLE_OPERATIONS["insertAfterBlock"]!(ids, "TableGrid"),
      ],
    });
    expect(result.applied).toEqual([]);
    expect(result.skipped).toContainEqual(expect.objectContaining({ reason: "missingStyle" }));
    expect(texts(reviewer)).toEqual(before);
  });
});

describe("a defined paragraph style still applies", () => {
  test.each(
    Object.entries(STYLE_OPERATIONS).flatMap(([operationName, factory]) =>
      supportedModes(factory).map((mode) => ({ operationName, factory, mode })),
    ),
  )("$operationName with Heading2, $mode mode", async ({ factory, mode }) => {
    const { reviewer, ids } = await openReviewer();
    const result = applyOne(reviewer, factory(ids, "Heading2"), mode);
    expect(result.skipped).toEqual([]);
    expect(result.applied.map(({ id }) => id)).toEqual(["op"]);
    if (mode === "suggested") {
      reviewer.acceptAll();
    }
    const saved = await reviewer.toBuffer();
    expect(await documentXml(saved)).toContain('<w:pStyle w:val="Heading2"/>');
  });

  test("setBlockParagraphProperties reopens as the heading it names", async () => {
    const { reviewer, ids } = await openReviewer();
    applyOne(reviewer, STYLE_OPERATIONS["setBlockParagraphProperties"]!(ids, "Heading2"), "direct");
    const reopened = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer(), {
      author: "AI",
    });
    expect(reopened.getContent().find((block) => block.text === ANCHOR)).toMatchObject({
      kind: "heading",
      headingLevel: 2,
    });
  });

  test("styleId: null still clears the style", async () => {
    const { reviewer, ids } = await openReviewer();
    const result = applyOne(
      reviewer,
      {
        id: "op",
        type: "setBlockParagraphProperties",
        blockId: ids.heading,
        properties: { styleId: null },
      },
      "direct",
    );
    expect(result.applied.map(({ id }) => id)).toEqual(["op"]);
    const xml = await documentXml(await reviewer.toBuffer());
    expect(xml).toContain(HEADING);
    expect(xml).not.toContain("<w:pStyle");
  });
});

describe("a caller copying another document's references keeps them", () => {
  test("undefinedStyles: keep writes the reference as given", async () => {
    const { reviewer, ids } = await openReviewer();
    const result = reviewer.applyDocumentOperationsToStory({
      story: { type: "main" },
      batch: {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "tracked-changes",
        operations: [STYLE_OPERATIONS["insertAfterBlock"]!(ids, "NoSuchStyle")],
      },
      undefinedStyles: "keep",
    });
    expect(result.applied.map(({ id }) => id)).toEqual(["op"]);
    expect(await documentXml(await reviewer.toBuffer())).toContain(
      '<w:pStyle w:val="NoSuchStyle"/>',
    );
  });
});
