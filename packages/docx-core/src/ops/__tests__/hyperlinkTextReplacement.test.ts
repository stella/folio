/**
 * Text replacing exactly a hyperlink's text keeps the link, edited directly
 * or tracked: the new text takes a copy of the link. A range reaching past
 * the link, or text typed at its outer edge, does not extend it. Accepting
 * the tracked replacement gives the direct edit; rejecting it gives the
 * source back.
 */

import { describe, expect, test } from "bun:test";

import type { Document, ParagraphContent } from "../../model/document";
import { applyDocumentOps } from "../apply";
import { normalizeForOps } from "../contract";
import { compileEditorIntent, type EditorIntentMode } from "../editorIntent";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  OP_STORIES,
  REVISION_DECISIONS,
} from "../types";

const LINK = {
  type: "hyperlink",
  rId: "rId9",
  href: "https://example.invalid/terms",
  tooltip: "Terms",
} as const;

const run = (text: string): ParagraphContent => ({
  type: "run",
  content: [{ type: "text", text }],
});

const source = normalizeForOps({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          paraId: "00000002",
          content: [
            run("See "),
            { ...LINK, children: [run("the standard terms")] },
            run(" now."),
          ],
        },
      ],
    },
  },
} satisfies Document);

const at = (offset: number) => ({
  story: OP_STORIES.MAIN,
  blockId: "00000002",
  offset,
});

const EDITING: EditorIntentMode = { type: "editing" };
const SUGGESTING: EditorIntentMode = {
  type: "suggesting",
  revision: { id: 2, author: "Reviewer", date: "2026-01-01T00:00:00Z" },
  newIds: { revision: [3, 4, 5, 6], control: [] },
};

/** Apply a plan, checking that its inverse restores the source exactly. */
const applied = (document: Document, ops: readonly DocumentOp[]) => {
  const result = applyDocumentOps(document, ops).unwrap();
  expect(
    applyDocumentOps(result.document, result.inverse).unwrap().document,
  ).toStrictEqual(document);
  return result.document;
};

const replaced = (
  { from, to, text }: { from: number; to: number; text: string },
  mode: EditorIntentMode,
) =>
  applied(
    source,
    compileEditorIntent(source, {
      intent: { type: "replaceText", from: at(from), to: at(to), text },
      mode,
    }).unwrap().ops,
  );

const paragraphContent = (document: Document) => {
  const [paragraph] = document.package.document.content;
  if (paragraph?.type !== "paragraph") throw new Error("Expected a paragraph");
  return paragraph.content;
};

/** The paragraph as text, each link's text in brackets. */
const reading = (document: Document): string => {
  const text = (node: unknown): string =>
    (JSON.stringify(node).match(/"text":"[^"]*"/gu) ?? [])
      .map((match) => match.slice(8, -1))
      .join("");
  const read = (node: ParagraphContent): string => {
    switch (node.type) {
      case "hyperlink":
        return `[${text(node)}]`;
      case "insertion":
      case "deletion":
        return node.content.map(read).join("");
      default:
        return text(node);
    }
  };
  return paragraphContent(document).map(read).join("");
};

const revisionIds = (document: Document): number[] =>
  paragraphContent(document).flatMap((node) =>
    node.type === "insertion" || node.type === "deletion" ? [node.info.id] : [],
  );

const resolved = (document: Document, decision: "accept" | "reject") =>
  applied(document, [
    {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds: revisionIds(document),
      decision:
        decision === "accept"
          ? REVISION_DECISIONS.ACCEPT
          : REVISION_DECISIONS.REJECT,
    },
  ]);

describe.each([
  {
    name: "exactly the link's text",
    edit: { from: 4, to: 22, text: "the revised terms" },
    direct: "See [the revised terms] now.",
  },
  {
    name: "a range reaching past the link",
    edit: { from: 3, to: 23, text: "X" },
    direct: "SeeXnow.",
  },
  {
    name: "a range reaching past the link's start",
    edit: { from: 3, to: 22, text: "X" },
    direct: "SeeX now.",
  },
  {
    name: "a caret at the link's end",
    edit: { from: 22, to: 22, text: "X" },
    direct: "See [the standard terms]X now.",
  },
  {
    name: "a caret at the link's start",
    edit: { from: 4, to: 4, text: "X" },
    direct: "See X[the standard terms] now.",
  },
])("replacing $name", ({ edit, direct }) => {
  test("directly gives the expected link", () => {
    expect(reading(replaced(edit, EDITING))).toBe(direct);
  });

  test("tracked, accepting gives the direct edit and rejecting the source", () => {
    const tracked = replaced(edit, SUGGESTING);
    expect(paragraphContent(resolved(tracked, "accept"))).toStrictEqual(
      paragraphContent(replaced(edit, EDITING)),
    );
    expect(resolved(tracked, "reject")).toStrictEqual(source);
  });
});

test("the kept link is a copy of the source link", () => {
  for (const mode of [EDITING, SUGGESTING]) {
    const links = JSON.stringify(
      paragraphContent(
        replaced({ from: 4, to: 22, text: "the revised terms" }, mode),
      ),
    );
    expect(links.match(/"rId":"rId9"/gu)?.length).toBe(
      mode.type === "editing" ? 1 : 2,
    );
    expect(links).toContain('"tooltip":"Terms"');
  }
});
