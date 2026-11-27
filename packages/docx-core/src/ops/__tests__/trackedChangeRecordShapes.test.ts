/**
 * Suggested changes over records a tracked change has to hold: a point
 * comment's reference in a deleted paragraph, a tracked change inside a
 * simple field's result, and a deleted paragraph mark that ends a section.
 * Each is recorded, accepted and rejected as its direct edit and its source
 * state are, and every step undoes.
 */
import { describe, expect, test } from "bun:test";

import type { Document, Paragraph } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { compileEditorIntent, type EditorIntent } from "../editorIntent";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  OP_STORIES,
  REVISION_DECISIONS,
  type RevisionDecision,
} from "../types";

const STAMP = { id: 100, author: "Reviewer", date: "2026-02-03T04:05:06Z" };
const NEW_IDS = {
  revision: Array.from({ length: 16 }, (_, index) => 101 + index),
};

const run = (text: string) => ({
  type: "run" as const,
  content: [{ type: "text" as const, text }],
});

const apply = (document: Document, ops: readonly DocumentOp[]) => {
  const result = applyDocumentOps(document, ops);
  if (result.isErr()) throw result.error;
  return result.value;
};

const compiled = (
  document: Document,
  intent: EditorIntent,
  suggesting: boolean,
): readonly DocumentOp[] => {
  const result = compileEditorIntent(document, {
    intent,
    mode: suggesting
      ? { type: "suggesting", revision: STAMP, newIds: NEW_IDS }
      : { type: "editing", newIds: NEW_IDS },
  });
  if (result.isErr()) throw result.error;
  return result.value.ops;
};

const resolve = (
  document: Document,
  revisionIds: readonly number[],
  decision: RevisionDecision,
) =>
  apply(document, [
    {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds,
      decision,
    },
  ]);

/**
 * The suggestion against its direct edit: accepting gives the direct result,
 * rejecting gives the source back, and the suggestion and each resolution
 * undo.
 */
const assertSuggestionLaws = (document: Document, intent: EditorIntent) => {
  const direct = apply(document, compiled(document, intent, false));
  const suggested = apply(document, compiled(document, intent, true));
  expect(suggested.revisions.length).toBeGreaterThan(0);
  expect(apply(suggested.document, suggested.inverse).document).toStrictEqual(
    document,
  );
  for (const decision of Object.values(REVISION_DECISIONS)) {
    const resolved = resolve(suggested.document, suggested.revisions, decision);
    expect(resolved.document).toStrictEqual(
      decision === REVISION_DECISIONS.ACCEPT ? direct.document : document,
    );
    expect(apply(resolved.document, resolved.inverse).document).toStrictEqual(
      suggested.document,
    );
  }
  return { direct, suggested };
};

describe("a point comment in a deleted paragraph", () => {
  const document: Document = {
    package: {
      document: {
        comments: [
          { id: 7, author: "Commenter", anchorKind: "point", content: [] },
        ],
        content: [
          {
            type: "paragraph",
            paraId: "00000001",
            content: [
              run("Noted "),
              { type: "commentReference", id: 7 },
              run("here."),
            ],
          },
          { type: "paragraph", paraId: "00000002", content: [run("Kept.")] },
        ],
      },
    },
  };
  const intent: EditorIntent = {
    type: "deleteParagraphs",
    story: OP_STORIES.MAIN,
    blockIds: ["00000001"],
  };

  test("the deletion takes the reference in; accepting retires the comment as the direct deletion does", () => {
    const { direct, suggested } = assertSuggestionLaws(document, intent);
    expect(direct.document.package.document.comments).toBeUndefined();
    const [deleted] = suggested.document.package.document.content;
    expect(deleted?.type === "paragraph" && deleted.content).toEqual([
      {
        type: "deletion",
        info: expect.objectContaining({ author: STAMP.author }),
        content: [
          run("Noted "),
          { type: "commentReference", id: 7 },
          run("here."),
        ],
      },
    ]);
    expect(suggested.document.package.document.comments).toEqual(
      document.package.document.comments,
    );
  });

  test("a raw tracked block deletion records it too", () => {
    const result = applyDocumentOp(document, {
      type: DOCUMENT_OP_TYPES.DELETE_BLOCKS,
      story: OP_STORIES.MAIN,
      blockIds: ["00000001"],
      revision: STAMP,
      newIds: NEW_IDS,
    });
    expect(result.isOk()).toBe(true);
  });
});

describe("a tracked change inside a simple field's result", () => {
  const info = (id: number) => ({
    id,
    author: "Earlier",
    date: "2026-01-01T00:00:00Z",
  });
  type FieldContent = Extract<
    Paragraph["content"][number],
    { type: "simpleField" }
  >["content"];
  const withField = (content: FieldContent): Document => ({
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            paraId: "00000001",
            content: [
              run("Page "),
              {
                type: "simpleField",
                instruction: "PAGE",
                fieldType: "PAGE",
                content,
              },
              run(" of the report."),
            ],
          },
        ],
      },
    },
  });
  const inserted = {
    type: "insertion" as const,
    info: info(5),
    content: [run("2")],
  };
  const deleted = {
    type: "deletion" as const,
    info: info(6),
    content: [run("3")],
  };
  const source = withField([run("1"), inserted, deleted]);

  test.each([
    [REVISION_DECISIONS.ACCEPT, [5], [run("1"), run("2"), deleted]],
    [REVISION_DECISIONS.REJECT, [5], [run("1"), deleted]],
    [REVISION_DECISIONS.ACCEPT, [6], [run("1"), inserted]],
    [REVISION_DECISIONS.REJECT, [6], [run("1"), inserted, run("3")]],
    [REVISION_DECISIONS.ACCEPT, [5, 6], [run("1"), run("2")]],
    [REVISION_DECISIONS.REJECT, [5, 6], [run("1"), run("3")]],
  ] as const)(
    "%s of %j resolves it inside the field",
    (decision, ids, shown) => {
      const resolved = resolve(source, ids, decision);
      expect(resolved.document).toStrictEqual(withField([...shown]));
      expect(apply(resolved.document, resolved.inverse).document).toStrictEqual(
        source,
      );
    },
  );

  test("a container the resolution empties goes; the field stays, showing nothing", () => {
    const linked = withField([
      {
        type: "inlineWrapper",
        kind: "smartTag",
        element: "place",
        content: [inserted],
      },
    ]);
    const rejected = resolve(linked, [5], REVISION_DECISIONS.REJECT);
    expect(rejected.document).toStrictEqual(withField([]));
    expect(apply(rejected.document, rejected.inverse).document).toStrictEqual(
      linked,
    );
  });
});

describe("a suggested join at a section break", () => {
  const document: Document = {
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            paraId: "00000001",
            sectionProperties: { pageWidth: 11_906 },
            content: [run("Ends a section.")],
          },
          {
            type: "paragraph",
            paraId: "00000002",
            content: [run("Opens the next.")],
          },
        ],
      },
    },
  };

  // Deleting a section break merges its section into the next one: the
  // joined paragraph keeps the later paragraph's mark and section, as the
  // direct join does.
  test.each<[string, EditorIntent]>([
    [
      "a join",
      {
        type: "joinParagraphs",
        story: OP_STORIES.MAIN,
        blockId: "00000001",
        nextBlockId: "00000002",
      },
    ],
    [
      "a deletion across the break",
      {
        type: "replaceText",
        from: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 4 },
        to: { story: OP_STORIES.MAIN, blockId: "00000002", offset: 5 },
        text: "",
      },
    ],
  ])(
    "%s records the deleted mark and resolves as the direct edit",
    (_shape, intent) => {
      const { direct, suggested } = assertSuggestionLaws(document, intent);
      const [first] = suggested.document.package.document.content;
      expect(first?.type === "paragraph" && first.pPrMark?.kind).toBe("del");
      expect(
        first?.type === "paragraph" && first.sectionProperties,
      ).toStrictEqual({ pageWidth: 11_906 });
      const [joined, ...rest] = direct.document.package.document.content;
      expect(rest).toEqual([]);
      expect(
        joined?.type === "paragraph" ? joined.sectionProperties : "none",
      ).toBeUndefined();
    },
  );
});
