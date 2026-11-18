import { describe, expect, test } from "bun:test";

import type { Document, Paragraph } from "../../model/document";
import { applyDocumentOps } from "../apply";
import { contractViolation } from "../contract";
import { allocateEditorIntentIds, compileEditorIntent } from "../editorIntent";
import type { EditorIntentMode } from "../editorIntent";
import type { HyperlinkEditorIntent } from "../hyperlinkIntent";
import {
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  REVISION_DECISIONS,
  type RevisionDecision,
} from "../types";

const paragraph = (id: string, text: string, styleId?: string): Paragraph => ({
  type: "paragraph",
  paraId: id,
  ...(styleId === undefined ? {} : { formatting: { styleId } }),
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

const original = (): Document => ({
  package: {
    document: {
      content: [
        paragraph("00000001", "Agreement", "Heading1"),
        paragraph("00000002", "Body text."),
        paragraph("00000003", "Closing."),
      ],
    },
  },
});

const at = (blockId: string, offset: number) => ({
  story: OP_STORIES.MAIN,
  blockId,
  offset,
});

const linkOver = (
  to: ReturnType<typeof at>,
): Extract<HyperlinkEditorIntent, { type: "insertHyperlink" }> => ({
  type: "insertHyperlink",
  from: at("00000001", 0),
  to,
  text: "a link",
  href: "https://example.org/",
});

const applied = (
  intent: HyperlinkEditorIntent,
  kind: "editing" | "suggesting",
) => {
  const document = original();
  const ids = allocateEditorIntentIds(document, intent);
  const mode: EditorIntentMode =
    kind === "editing"
      ? { type: "editing", newIds: ids.newIds }
      : {
          type: "suggesting",
          revision: {
            id: ids.revisionId,
            author: "Reviewer",
            date: "2026-11-18T00:00:00Z",
          },
          newIds: ids.newIds,
        };
  const compiled = compileEditorIntent(document, { intent, mode }).unwrap();
  const result = applyDocumentOps(document, compiled.ops).unwrap();
  expect(contractViolation(result.document)).toBeUndefined();
  expect(
    applyDocumentOps(result.document, result.inverse).unwrap().document,
  ).toStrictEqual(document);
  return result;
};

const resolved = (
  result: ReturnType<typeof applied>,
  decision: RevisionDecision,
): Document =>
  applyDocumentOps(result.document, [
    {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds: result.revisions,
      decision,
    },
  ]).unwrap().document;

/** Each paragraph's style and visible text, links as their targets. */
const outline = (document: Document) =>
  document.package.document.content.map((block) =>
    block.type === "paragraph"
      ? {
          styleId: block.formatting?.styleId,
          content: block.content.map((item) =>
            item.type === "hyperlink" ? `link:${item.href ?? ""}` : item.type,
          ),
        }
      : block.type,
  );

describe("a link inserted over paragraphs", () => {
  for (const [name, to] of [
    ["the whole document", at("00000003", 8)],
    ["the first paragraphs", at("00000002", 4)],
  ] as const) {
    test(`over ${name} keeps the first paragraph's properties, made directly or suggested`, () => {
      const intent = linkOver(to);
      const direct = applied(intent, "editing").document;
      const [first] = outline(direct);
      expect(first).toMatchObject({ styleId: "Heading1" });
      const suggested = applied(intent, "suggesting");
      expect(outline(resolved(suggested, REVISION_DECISIONS.ACCEPT))).toEqual(
        outline(direct),
      );
      expect(outline(resolved(suggested, REVISION_DECISIONS.REJECT))).toEqual(
        outline(original()),
      );
    });
  }
});
