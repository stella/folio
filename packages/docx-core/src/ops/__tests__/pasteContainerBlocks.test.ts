/**
 * Clipboard block content controls, custom XML elements and kept block markup:
 * pasted directly, each stands whole before the caret's paragraph with its
 * own paragraphs and control identified afresh, and the operations' inverse
 * gives the document back. A suggestion has no record for them and refuses.
 */
import { describe, expect, test } from "bun:test";

import type {
  BlockContent,
  Document,
  Paragraph,
  Run,
} from "../../model/document";
import { applyDocumentOps } from "../apply";
import { validateOpsDocument } from "../contract";
import { compileEditorIntent, type EditorIntentMode } from "../editorIntent";
import type { PastedBlock } from "../pastedBlocks";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import { OP_STORIES } from "../types";

const run = (text: string): Run => ({
  type: "run",
  content: [{ type: "text", text }],
});
const paragraph = (text: string, paraId?: string): Paragraph => ({
  type: "paragraph",
  ...(paraId === undefined ? {} : { paraId }),
  content: text === "" ? [] : [run(text)],
});

const source = (): Document => ({
  package: {
    document: {
      content: [
        {
          type: "blockSdt",
          properties: { sdtType: "richText", id: 7 },
          content: [paragraph("held", "00000003")],
        },
        paragraph("before after", "00000001"),
        paragraph("next", "00000002"),
      ],
    },
  },
});

const control = (): PastedBlock => ({
  type: "blockSdt",
  properties: { sdtType: "richText", id: 7, tag: "clause" },
  content: [paragraph("Inside one", "00000003"), paragraph("Inside two")],
});
const customXml = (): PastedBlock => ({
  type: "blockCustomXml",
  openingXml: '<w:customXml w:element="clause">',
  closingXml: "</w:customXml>",
  content: [
    {
      type: "blockSdt",
      properties: { sdtType: "richText", id: 8 },
      content: [paragraph("Nested")],
    },
  ],
});
const preserved = (): PastedBlock => ({
  type: "preservedBlock",
  xml: '<w:permStart w:id="1"/>',
});

const DIRECT = { type: "editing" } as const satisfies EditorIntentMode;

const pasted = (blocks: readonly PastedBlock[], mode: EditorIntentMode) => {
  const document = source();
  const at = { story: OP_STORIES.MAIN, blockId: "00000001", offset: 7 };
  return {
    document,
    compiled: compileEditorIntent(document, {
      intent: {
        type: "pasteBlocks",
        from: at,
        to: at,
        blocks: [paragraph("Lead"), ...blocks, paragraph("Trail")],
        openStart: 1,
        openEnd: 1,
      },
      mode,
    }),
  };
};

const paraIds = (blocks: readonly BlockContent[]): string[] =>
  blocks.flatMap((block) => {
    if (block.type === "paragraph")
      return block.paraId === undefined ? [] : [block.paraId];
    if (block.type === "blockSdt" || block.type === "blockCustomXml")
      return paraIds(block.content);
    return [];
  });

describe.each([
  { kind: "block content control", blocks: [control()] },
  { kind: "custom XML element", blocks: [customXml()] },
  { kind: "kept block markup", blocks: [preserved()] },
  { kind: "every kind", blocks: [control(), customXml(), preserved()] },
])("pasting a $kind", ({ blocks }) => {
  test("places it whole between the split halves, and undoes", () => {
    const { document, compiled } = pasted(blocks, DIRECT);
    if (compiled.isErr()) throw compiled.error;
    const applied = applyDocumentOps(document, compiled.value.ops);
    if (applied.isErr()) throw applied.error;
    const content = applied.value.document.package.document.content;
    expect(validateOpsDocument(applied.value.document).isOk()).toBe(true);
    expect(content.map(({ type }) => type)).toEqual([
      "blockSdt",
      "paragraph",
      ...blocks.map(({ type }) => type),
      "paragraph",
      "paragraph",
    ]);
    // Every paragraph id is the package's once; no control keeps a copied id.
    const ids = paraIds(content);
    expect(new Set(ids).size).toBe(ids.length);
    const controlIds: number[] = [];
    const visit = (items: readonly BlockContent[]) => {
      for (const item of items) {
        if (item.type === "blockSdt") {
          if (item.properties.id !== undefined)
            controlIds.push(item.properties.id);
          visit(item.content);
        }
        if (item.type === "blockCustomXml") visit(item.content);
      }
    };
    visit(content);
    expect(new Set(controlIds).size).toBe(controlIds.length);
    const undone = applyDocumentOps(
      applied.value.document,
      applied.value.inverse,
    );
    if (undone.isErr()) throw undone.error;
    expect(undone.value.document.package.document.content).toEqual(
      document.package.document.content,
    );
  });

  test("refuses as a suggestion, which has no record for it", () => {
    const { compiled } = pasted(blocks, {
      type: "suggesting",
      revision: { id: 90, author: "R", date: "2026-01-02T00:00:00Z" },
      newIds: {
        revision: Array.from({ length: 80 }, (_, index) => 91 + index),
      },
    });
    expect(compiled.isErr()).toBe(true);
    if (compiled.isErr())
      expect(compiled.error.reason).toBe(
        DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
      );
  });
});

describe("dragging a block content control", () => {
  const document = (): Document => ({
    package: {
      document: {
        content: [
          paragraph("zero", "00000001"),
          paragraph("one", "00000002"),
          {
            type: "blockSdt",
            properties: { sdtType: "richText", id: 7 },
            content: [paragraph("held", "00000005")],
          },
          paragraph("two", "00000003"),
          paragraph("three", "00000004"),
        ],
      },
    },
  });
  const shapeOf = (blocks: readonly BlockContent[]): unknown[] =>
    blocks.map((block) => {
      if (block.type === "blockSdt") return { control: shapeOf(block.content) };
      if (block.type !== "paragraph") return block.type;
      return block.content
        .flatMap((item) =>
          item.type === "run"
            ? item.content.flatMap((leaf) =>
                leaf.type === "text" ? [leaf.text] : [],
              )
            : [],
        )
        .join("");
    });
  const drag = (mode: EditorIntentMode) => {
    const dragged = document();
    const at = (blockId: string, offset: number) => ({
      story: OP_STORIES.MAIN,
      blockId,
      offset,
    });
    return {
      dragged,
      compiled: compileEditorIntent(dragged, {
        intent: {
          type: "moveBlocks",
          from: at("00000002", 3),
          to: at("00000003", 0),
          target: at("00000004", 2),
          blocks: [
            paragraph(""),
            {
              type: "blockSdt",
              properties: { sdtType: "richText", id: 7 },
              content: [paragraph("held")],
            },
            paragraph(""),
          ],
          openStart: 1,
          openEnd: 1,
        },
        mode,
      }),
    };
  };

  test("moves it directly, and the inverse puts it back", () => {
    const { dragged, compiled } = drag(DIRECT);
    if (compiled.isErr()) throw compiled.error;
    const applied = applyDocumentOps(dragged, compiled.value.ops);
    if (applied.isErr()) throw applied.error;
    expect(validateOpsDocument(applied.value.document).isOk()).toBe(true);
    expect(shapeOf(applied.value.document.package.document.content)).toEqual([
      "zero",
      "onetwo",
      "th",
      { control: ["held"] },
      "ree",
    ]);
    const undone = applyDocumentOps(
      applied.value.document,
      applied.value.inverse,
    );
    if (undone.isErr()) throw undone.error;
    expect(undone.value.document).toEqual(document());
  });

  test("refuses as a suggestion, which has no record for it", () => {
    const { compiled } = drag({
      type: "suggesting",
      revision: { id: 90, author: "R", date: "2026-01-02T00:00:00Z" },
      newIds: {
        revision: Array.from({ length: 80 }, (_, index) => 91 + index),
      },
    });
    expect(compiled.isErr()).toBe(true);
    if (compiled.isErr())
      expect(compiled.error.reason).toBe(
        DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
      );
  });
});
