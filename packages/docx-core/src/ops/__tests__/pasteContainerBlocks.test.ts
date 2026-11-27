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

// A block-level `w:sdt` takes no revision of its own (CT_SdtBlock records
// none), so OOXML has no record for its suggested paste or move. A block
// `w:customXml` has range records the model does not carry yet, and kept
// markup is not read.
const NO_RECORD = DOCUMENT_OP_REFUSAL_REASONS.NO_REVISION_RECORD;
const NOT_YET = DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE;

describe.each([
  { kind: "block content control", blocks: [control()], suggested: NO_RECORD },
  { kind: "custom XML element", blocks: [customXml()], suggested: NOT_YET },
  { kind: "kept block markup", blocks: [preserved()], suggested: NOT_YET },
  {
    kind: "every kind",
    blocks: [control(), customXml(), preserved()],
    suggested: NO_RECORD,
  },
])("pasting a $kind", ({ blocks, suggested }) => {
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
    if (compiled.isErr()) expect(compiled.error.reason).toBe(suggested);
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
    if (compiled.isErr()) expect(compiled.error.reason).toBe(NO_RECORD);
  });
});

describe("a pasted container's paragraph identities", () => {
  test("pass over the ids the session retired", () => {
    const document: Document = {
      package: { document: { content: [paragraph("next", "00000001")] } },
    };
    const at = { story: OP_STORIES.MAIN, blockId: "00000001", offset: 0 };
    const compiled = compileEditorIntent(document, {
      intent: {
        type: "pasteBlocks",
        from: at,
        to: at,
        blocks: [
          {
            type: "blockSdt",
            properties: { sdtType: "richText", id: 7 },
            content: [paragraph("held")],
          },
        ],
        openStart: 0,
        openEnd: 0,
      },
      mode: { type: "editing", reservedBlockIds: new Set(["00000002"]) },
    });
    if (compiled.isErr()) throw compiled.error;
    const applied = applyDocumentOps(document, compiled.value.ops);
    if (applied.isErr()) throw applied.error;
    const ids = paraIds(applied.value.document.package.document.content);
    expect(ids).not.toContain("00000002");
    expect(new Set(ids).size).toBe(ids.length);
    expect(validateOpsDocument(applied.value.document).isOk()).toBe(true);
  });
});

const inControl = (held: Paragraph): PastedBlock => ({
  type: "blockSdt",
  properties: { sdtType: "richText", id: 7 },
  content: [held],
});
const inTable = (held: Paragraph): PastedBlock => ({
  type: "table",
  rows: [{ type: "tableRow", cells: [{ type: "tableCell", content: [held] }] }],
});

describe.each([
  {
    kind: "block content control",
    wrap: inControl,
    suggested: NO_RECORD,
  },
  // A suggestion keeps the source's reference; its copy would need a note.
  {
    kind: "table",
    wrap: inTable,
    suggested: DOCUMENT_OP_REFUSAL_REASONS.INVALID_OPERATION,
  },
])(
  "dragging a $kind holding a note reference and a comment",
  ({ wrap, suggested }) => {
    const held = (): Paragraph => ({
      type: "paragraph",
      paraId: "00000005",
      content: [
        { type: "commentRangeStart", id: 4 },
        run("held"),
        { type: "run", content: [{ type: "footnoteRef", id: 1 }] },
        { type: "commentRangeEnd", id: 4 },
        { type: "commentReference", id: 4 },
      ],
    });
    const document = (): Document => ({
      package: {
        document: {
          content: [
            paragraph("one", "00000002"),
            wrap(held()),
            paragraph("two", "00000003"),
            paragraph("three", "00000004"),
          ],
          comments: [
            {
              id: 4,
              author: "Reviewer",
              date: "2026-01-02T00:00:00Z",
              content: [paragraph("Standing comment", "0000C004")],
            },
          ],
        },
        footnotes: [
          {
            type: "footnote",
            id: 1,
            content: [paragraph("Standing note", "0000F001")],
          },
        ],
      },
    });
    const at = (blockId: string, offset: number) => ({
      story: OP_STORIES.MAIN,
      blockId,
      offset,
    });
    const drag = (dragged: Document, mode: EditorIntentMode) => {
      const { paraId: _paraId, ...copied } = held();
      return compileEditorIntent(dragged, {
        intent: {
          type: "moveBlocks",
          from: at("00000002", 3),
          to: at("00000003", 0),
          target: at("00000004", 2),
          blocks: [paragraph(""), wrap(copied), paragraph("")],
          openStart: 1,
          openEnd: 1,
        },
        mode,
      });
    };
    const anchorsIn = (moved: Document) => {
      const text = JSON.stringify(moved.package.document.content);
      return {
        notes: text.match(/"footnoteRef","id":1/gu)?.length ?? 0,
        comments:
          text.match(/"comment(?:RangeStart|RangeEnd|Reference)","id":4/gu)
            ?.length ?? 0,
      };
    };

    test("moves it directly, keeping the note and the comment under their ids", () => {
      const dragged = document();
      const compiled = drag(dragged, DIRECT);
      if (compiled.isErr()) throw compiled.error;
      const applied = applyDocumentOps(dragged, compiled.value.ops);
      if (applied.isErr()) throw applied.error;
      const moved = applied.value.document;
      expect(validateOpsDocument(moved).isOk()).toBe(true);
      expect(moved.package.document.content.map(({ type }) => type)).toEqual([
        "paragraph",
        "paragraph",
        wrap(held()).type,
        "paragraph",
      ]);
      // One reference and one set of markers, each naming the one it named.
      expect(anchorsIn(moved)).toEqual({ notes: 1, comments: 3 });
      expect(moved.package.footnotes).toEqual(document().package.footnotes);
      expect(moved.package.document.comments).toEqual(
        document().package.document.comments,
      );
      const undone = applyDocumentOps(moved, applied.value.inverse);
      if (undone.isErr()) throw undone.error;
      expect(undone.value.document).toEqual(document());
    });

    test("refuses as a suggestion", () => {
      const compiled = drag(document(), {
        type: "suggesting",
        revision: { id: 90, author: "R", date: "2026-01-02T00:00:00Z" },
        newIds: {
          revision: Array.from({ length: 80 }, (_, index) => 91 + index),
        },
      });
      expect(compiled.isErr()).toBe(true);
      if (compiled.isErr()) expect(compiled.error.reason).toBe(suggested);
    });
  },
);
