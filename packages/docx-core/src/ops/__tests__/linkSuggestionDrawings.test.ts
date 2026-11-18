import { describe, expect, test } from "bun:test";

import type {
  Document,
  ParagraphContent,
  RunContent,
} from "../../model/document";
import { applyDocumentOps } from "../apply";
import { contractViolation } from "../contract";
import { allocateEditorIntentIds, compileEditorIntent } from "../editorIntent";
import type { HyperlinkEditorIntent } from "../hyperlinkIntent";
import { paragraphIdsIn } from "../ids";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import {
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  REVISION_DECISIONS,
  type RevisionDecision,
} from "../types";
import { RUN_LEAVES } from "./runLeafSamples";

const PARA_ID = "12345678";
const TEXT_BOX_PARA_ID = "0000ABCD";

const textBoxShape: RunContent = {
  type: "shape",
  shape: {
    type: "shape",
    shapeType: "rect",
    size: { width: 10, height: 10 },
    textBody: {
      content: [
        {
          type: "paragraph",
          paraId: TEXT_BOX_PARA_ID,
          content: [{ type: "run", content: [{ type: "text", text: "box" }] }],
        },
      ],
    },
  },
};

const documentOf = (leaf: RunContent): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          paraId: PARA_ID,
          content: [
            { type: "run", content: [{ type: "text", text: "ab" }] },
            { type: "run", content: [leaf] },
            { type: "run", content: [{ type: "text", text: "cd" }] },
          ],
        },
      ],
    },
  },
});

const at = (offset: number) => ({
  story: OP_STORIES.MAIN,
  blockId: PARA_ID,
  offset,
});

const intent: HyperlinkEditorIntent = {
  type: "setHyperlink",
  from: at(0),
  to: at(5),
  href: "https://example.org/",
};

const paragraphContent = (document: Document): ParagraphContent[] => {
  const [paragraph] = document.package.document.content;
  return paragraph?.type === "paragraph" ? paragraph.content : [];
};

const suggested = (document: Document) => {
  const ids = allocateEditorIntentIds(document, intent);
  return compileEditorIntent(document, {
    intent,
    mode: {
      type: "suggesting",
      revision: {
        id: ids.revisionId,
        author: "Reviewer",
        date: "2026-11-18T00:00:00Z",
      },
      newIds: ids.newIds,
    },
  });
};

const direct = (document: Document): Document => {
  const ids = allocateEditorIntentIds(document, intent);
  const compiled = compileEditorIntent(document, {
    intent,
    mode: { type: "editing", newIds: ids.newIds },
  }).unwrap();
  return applyDocumentOps(document, compiled.ops).unwrap().document;
};

/** The paragraph's shape: links and runs, with each run's leaf kinds. */
const outline = (content: readonly ParagraphContent[]): unknown =>
  content.map((item) => {
    if (item.type === "hyperlink")
      return { link: item.href, children: outline(item.children) };
    if (item.type !== "run") return item.type;
    return item.content.map((leaf) =>
      leaf.type === "text" ? leaf.text : leaf.type,
    );
  });

describe("a link suggestion over a drawing", () => {
  for (const [name, leaf] of [
    ["picture", RUN_LEAVES.drawing],
    ["shape", RUN_LEAVES.shape],
    ["text box", textBoxShape],
  ] as const) {
    test(`strikes the ${name} and inserts a linked copy that accepts to the direct edit`, () => {
      const document = documentOf(leaf);
      const compiled = suggested(document).unwrap();
      const applied = applyDocumentOps(document, compiled.ops).unwrap();
      expect(contractViolation(applied.document)).toBeUndefined();
      // Undo and redo are exact.
      const undone = applyDocumentOps(
        applied.document,
        applied.inverse,
      ).unwrap();
      expect(undone.document).toStrictEqual(document);
      const resolve = (decision: RevisionDecision) =>
        applyDocumentOps(applied.document, [
          {
            type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
            story: OP_STORIES.MAIN,
            revisionIds: applied.revisions,
            decision,
          },
        ]).unwrap().document;
      const accepted = resolve(REVISION_DECISIONS.ACCEPT);
      expect(outline(paragraphContent(accepted))).toEqual(
        outline(paragraphContent(direct(document))),
      );
      expect(
        outline(paragraphContent(resolve(REVISION_DECISIONS.REJECT))),
      ).toEqual(outline(paragraphContent(document)));
      // Paragraphs inside the copy take ids of their own.
      const ids = paragraphIdsIn(applied.document.package);
      expect(new Set(ids).size).toBe(ids.length);
      if (leaf === textBoxShape) expect(ids).toHaveLength(3);
    });
  }

  test("refuses a copy of a text box holding a bookmark, which names a record of its own", () => {
    const shape = structuredClone(textBoxShape);
    if (shape.type !== "shape") throw new Error("expected a shape");
    const paragraph = shape.shape.textBody?.content[0];
    if (paragraph?.type !== "paragraph")
      throw new Error("expected a paragraph");
    paragraph.content.unshift(
      { type: "bookmarkStart", id: 3, name: "inside" },
      { type: "bookmarkEnd", id: 3 },
    );
    const compiled = suggested(documentOf(shape));
    expect(compiled.isErr() && compiled.error.reason).toBe(
      DOCUMENT_OP_REFUSAL_REASONS.UNRECORDABLE_LINK_SUGGESTION,
    );
  });
});
