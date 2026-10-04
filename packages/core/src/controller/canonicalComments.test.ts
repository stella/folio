import { expect, test, setDefaultTimeout } from "bun:test";
import fc from "fast-check";
import { EditorState } from "prosemirror-state";
import { freshCommentId } from "@stll/docx-core/ops";
import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { canonicalCommentBody, compileCanonicalComments } from "./canonicalComments";
import { createCanonicalSession, publishCanonicalProjection } from "./canonicalSession";
import type { Document } from "../types/document";
import { cloneDocumentWithParagraphPropertySources } from "../docx/paragraphPropertySource";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import { describePackageDifferences } from "../../../../scripts/lib/corpus-invariants/model-equality";

setDefaultTimeout(propertyTestTimeout(30_000));

const seed = () =>
  ({
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            paraId: "ABCDEF01",
            content: [{ type: "run", content: [{ type: "text", text: "A😀éB" }] }],
          },
        ],
      },
    },
  }) satisfies Document;

const setup = (document = seed()) => {
  const session = createCanonicalSession(document).unwrap();
  let state = EditorState.create({ doc: session.projection.doc });
  const apply = (command: Parameters<typeof compileCanonicalComments>[0]["command"]) => {
    const compiled = compileCanonicalComments({ document: session.document, command }).unwrap();
    if (compiled.ops.length === 0) return;
    const prepared = session.prepareOperations(state, compiled.ops).unwrap();
    state = publishCanonicalProjection({ state, commit: prepared, session }).unwrap().state;
  };
  const history = (direction: "undo" | "redo") => {
    const commit = direction === "undo" ? session.prepareUndo(state) : session.prepareRedo(state);
    state = publishCanonicalProjection({ state, commit: commit.unwrap(), session }).unwrap().state;
  };
  const create = (text: string, kind: "point" | "range") => {
    const id = freshCommentId(session.document).unwrap();
    const paragraph = session.document.package.document.content.at(0);
    if (paragraph?.type !== "paragraph" || paragraph.paraId === undefined)
      throw new Error("Missing canonical paragraph");
    const at = { story: "main", blockId: paragraph.paraId, offset: 1 } as const;
    apply({
      type: "create",
      comment: {
        id,
        author: "reviewer",
        done: false,
        date: "2026-01-01T00:00:00Z",
        content: canonicalCommentBody(session.document, text),
      },
      anchor: kind === "point" ? { kind, at } : { kind, from: at, to: { ...at, offset: 5 } },
    });
    return id;
  };
  return { session, apply, history, create };
};

test("generated point/range comment histories preserve exact undo, projection and modeled save", async () => {
  await assertProperty(
    fc.asyncProperty(
      fc.array(fc.stringMatching(/^[a-z]{1,12}$/u), { minLength: 1, maxLength: 6 }),
      fc.constantFrom("point", "range"),
      async (texts, kind) => {
        // Parsing the seed gives the model/transport oracle one shared baseline.
        const editor = setup(await parseDocx(await createDocx(seed())));
        const original = editor.session.document;
        let steps = 0;
        for (const text of texts) {
          const id = editor.create(text, kind);
          steps += 1;
          const replyId = freshCommentId(editor.session.document).unwrap();
          editor.apply({
            type: "create",
            comment: {
              id: replyId,
              author: "reply",
              done: false,
              content: canonicalCommentBody(editor.session.document, text + "reply"),
            },
            anchor: { kind: "reply", parentId: id },
          });
          steps += 1;
          const content = editor.session.document.package.document.comments?.find(
            (comment) => comment.id === id,
          )?.content;
          expect(content).toBeDefined();
          editor.apply({
            type: "update",
            id,
            content:
              content?.map((paragraph) =>
                Object.assign({}, paragraph, {
                  content: [
                    { type: "run", content: [{ type: "text", text: text + "edited" }] },
                  ] satisfies typeof paragraph.content,
                }),
              ) ?? [],
          });
          steps += 1;
          editor.apply({ type: "resolve", id, status: "resolved" });
          steps += 1;
          const committed = cloneDocumentWithParagraphPropertySources(editor.session.document);
          const reopened = await parseDocx(
            await createDocx(cloneDocumentWithParagraphPropertySources(committed)),
          );
          const differences = describePackageDifferences(committed, reopened);
          expect(differences).toEqual({ messages: [], omitted: 0 });
          editor.apply({ type: "delete", id });
          expect(
            editor.session.document.package.document.comments?.some(
              (comment) => comment.id === id || comment.id === replyId,
            ),
          ).toBe(false);
          editor.history("undo");
          expect(editor.session.document).toEqual(committed);
          editor.history("redo");
          const deleted = cloneDocumentWithParagraphPropertySources(editor.session.document);
          const reopenedDeleted = await parseDocx(
            await createDocx(cloneDocumentWithParagraphPropertySources(deleted)),
          );
          expect(describePackageDifferences(deleted, reopenedDeleted)).toEqual({
            messages: [],
            omitted: 0,
          });
          steps += 1;
        }
        const final = editor.session.document;
        for (let index = 0; index < steps; index += 1) editor.history("undo");
        expect(editor.session.document).toEqual(original);
        for (let index = 0; index < steps; index += 1) editor.history("redo");
        expect(editor.session.document).toEqual(final);
        expect(editor.session.projection.doc.textContent).toBe("A😀éB");
      },
    ),
    { numRuns: 12, seed: 20261005 },
  );
});

test("controlled comment echoes are inert; content/status edits compile and unanchored roots refuse atomically", () => {
  const editor = setup();
  const id = editor.create("note", "range");
  const comments = editor.session.document.package.document.comments ?? [];
  const echo = compileCanonicalComments({
    document: editor.session.document,
    command: { type: "replace", comments: structuredClone(comments) },
  }).unwrap();
  expect(echo.ops).toEqual([]);
  const next = comments.map((comment) => ({
    ...comment,
    done: true,
    content: comment.content.map((paragraph) => ({
      ...paragraph,
      content: [{ type: "run" as const, content: [{ type: "text" as const, text: "changed" }] }],
    })),
  }));
  editor.apply({ type: "replace", comments: next });
  expect(editor.session.document.package.document.comments?.at(0)?.done).toBe(true);
  editor.history("undo");
  expect(editor.session.document.package.document.comments).toEqual(comments);
  const before = editor.session.document;
  const refused = compileCanonicalComments({
    document: before,
    command: {
      type: "replace",
      comments: [
        ...comments,
        { id: id + 100, author: "host", content: canonicalCommentBody(before, "unanchored") },
      ],
    },
  });
  expect(refused.isErr()).toBe(true);
  expect(editor.session.document).toBe(before);
});

test("controlled replies compile parent-first and cannot retain an orphan", () => {
  const editor = setup();
  const rootId = editor.create("root", "range");
  const root = editor.session.document.package.document.comments?.at(0);
  expect(root).toBeDefined();
  if (root === undefined) throw new Error("Missing root");
  const firstId = freshCommentId(editor.session.document).unwrap();
  const parent = {
    id: firstId,
    author: "reply",
    done: false,
    parentId: rootId,
    content: canonicalCommentBody(editor.session.document, "parent"),
  };
  const withParent = compileCanonicalComments({
    document: editor.session.document,
    command: {
      type: "create",
      comment: { id: firstId, author: "reply", done: false, content: parent.content },
      anchor: { kind: "reply", parentId: rootId },
    },
  }).unwrap().document;
  const child = {
    id: freshCommentId(withParent).unwrap(),
    author: "reply",
    done: false,
    parentId: firstId,
    content: canonicalCommentBody(withParent, "child"),
  };
  editor.apply({ type: "replace", comments: [child, parent, root] });
  const committed = editor.session.document;
  expect(committed.package.document.comments?.map(({ id }) => id)).toEqual([
    rootId,
    firstId,
    child.id,
  ]);
  const orphan = compileCanonicalComments({
    document: committed,
    command: { type: "replace", comments: [child, parent] },
  });
  expect(orphan.isErr()).toBe(true);
  editor.history("undo");
  expect(editor.session.document.package.document.comments).toEqual([root]);
  editor.history("redo");
  expect(editor.session.document).toEqual(committed);
});

test("creating then deleting a canonical thread saves an owned empty comments part", async () => {
  const editor = setup(await parseDocx(await createDocx(seed())));
  const id = editor.create("temporary", "range");
  editor.apply({ type: "delete", id });
  const committed = cloneDocumentWithParagraphPropertySources(editor.session.document);
  const reopened = await parseDocx(
    await createDocx(cloneDocumentWithParagraphPropertySources(committed)),
  );
  expect(describePackageDifferences(committed, reopened)).toEqual({ messages: [], omitted: 0 });
});
