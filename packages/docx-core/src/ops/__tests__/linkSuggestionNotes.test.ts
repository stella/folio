import { describe, expect, test } from "bun:test";

import type {
  Document,
  Endnote,
  Footnote,
  ParagraphContent,
} from "../../model/document";
import { applyDocumentOps } from "../apply";
import { contractViolation } from "../contract";
import { allocateEditorIntentIds, compileEditorIntent } from "../editorIntent";
import type { HyperlinkEditorIntent } from "../hyperlinkIntent";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import {
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  REVISION_DECISIONS,
  type RevisionDecision,
} from "../types";

const PARA_ID = "12345678";

const note = <Kind extends "footnote" | "endnote">(
  kind: Kind,
  extra: ParagraphContent[] = [],
) =>
  ({
    type: kind,
    id: 1,
    content: [
      {
        type: "paragraph",
        paraId: kind === "footnote" ? "0000F001" : "0000E001",
        content: [
          { type: "run", content: [{ type: "noteMarker", kind }] },
          ...extra,
          { type: "run", content: [{ type: "text", text: " The note." }] },
        ],
      },
    ],
  }) as Kind extends "footnote" ? Footnote : Endnote;

const documentOf = (
  kind: "footnote" | "endnote",
  extra: ParagraphContent[] = [],
): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          paraId: PARA_ID,
          content: [
            { type: "run", content: [{ type: "text", text: "ab" }] },
            {
              type: "run",
              formatting: { styleId: "FootnoteReference" },
              content: [
                kind === "footnote"
                  ? { type: "footnoteRef", id: 1 }
                  : { type: "endnoteRef", id: 1 },
              ],
            },
            { type: "run", content: [{ type: "text", text: "cd" }] },
          ],
        },
      ],
    },
    ...(kind === "footnote"
      ? { footnotes: [note("footnote", extra)] }
      : { endnotes: [note("endnote", extra)] }),
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

const compiled = (document: Document, kind: "editing" | "suggesting") => {
  const ids = allocateEditorIntentIds(document, intent);
  return compileEditorIntent(document, {
    intent,
    mode:
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
          },
  });
};

/** The paragraph's links and runs, each run's leaves by text or kind. */
const outline = (content: readonly ParagraphContent[]): unknown =>
  content.map((item) => {
    if (item.type === "hyperlink")
      return { link: item.href, children: outline(item.children) };
    if (item.type !== "run") return item.type;
    return item.content.map((leaf) =>
      leaf.type === "text" ? leaf.text : leaf.type,
    );
  });

const body = (document: Document) => {
  const [paragraph] = document.package.document.content;
  return paragraph?.type === "paragraph" ? paragraph.content : [];
};

const notesOf = (document: Document, kind: "footnote" | "endnote") =>
  (kind === "footnote"
    ? document.package.footnotes
    : document.package.endnotes) ?? [];

/** A note's text, which a copy repeats under its own id. */
const noteText = (value: Footnote | Endnote) =>
  JSON.stringify(
    value.content.map((block) =>
      block.type === "paragraph" ? block.content : block,
    ),
  );

describe("a link suggestion over a note reference", () => {
  for (const kind of ["footnote", "endnote"] as const) {
    test(`strikes the ${kind} reference and inserts a linked copy naming a copy of its note`, () => {
      const document = documentOf(kind);
      const suggestion = compiled(document, "suggesting").unwrap();
      const applied = applyDocumentOps(document, suggestion.ops).unwrap();
      expect(contractViolation(applied.document)).toBeUndefined();
      const pending = notesOf(applied.document, kind);
      expect(pending.map(({ id }) => id)).toEqual([1, 2]);
      const [original, copy] = pending;
      if (original === undefined || copy === undefined)
        throw new Error("expected both notes");
      expect(noteText(copy)).toBe(noteText(original));
      const paraIds = applied.document.package.document.content
        .concat(pending.flatMap(({ content }) => content))
        .map((block) => (block.type === "paragraph" ? block.paraId : ""));
      expect(new Set(paraIds).size).toBe(paraIds.length);
      // Undo and redo are exact.
      const undone = applyDocumentOps(
        applied.document,
        applied.inverse,
      ).unwrap();
      expect(undone.document).toStrictEqual(document);
      expect(
        applyDocumentOps(undone.document, undone.inverse).unwrap().document,
      ).toStrictEqual(applied.document);

      const resolve = (decision: RevisionDecision) =>
        applyDocumentOps(applied.document, [
          {
            type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
            story: OP_STORIES.MAIN,
            revisionIds: applied.revisions,
            decision,
          },
        ]).unwrap().document;
      const direct = applyDocumentOps(
        document,
        compiled(document, "editing").unwrap().ops,
      ).unwrap().document;
      const accepted = resolve(REVISION_DECISIONS.ACCEPT);
      expect(outline(body(accepted))).toEqual(outline(body(direct)));
      // The copy's note stays; the struck reference's goes with it.
      expect(notesOf(accepted, kind).map(noteText)).toEqual(
        notesOf(direct, kind).map(noteText),
      );
      expect(notesOf(accepted, kind)).toHaveLength(1);
      const rejected = resolve(REVISION_DECISIONS.REJECT);
      expect(outline(body(rejected))).toEqual(outline(body(document)));
      expect(notesOf(rejected, kind)).toStrictEqual(notesOf(document, kind));
    });
  }

  test("refuses to copy a note holding a bookmark, which names a record of its own", () => {
    const result = compiled(
      documentOf("footnote", [
        { type: "bookmarkStart", id: 4, name: "inNote" },
        { type: "bookmarkEnd", id: 4 },
      ]),
      "suggesting",
    );
    expect(result.isErr() && result.error.reason).toBe(
      DOCUMENT_OP_REFUSAL_REASONS.UNRECORDABLE_LINK_SUGGESTION,
    );
  });
});
