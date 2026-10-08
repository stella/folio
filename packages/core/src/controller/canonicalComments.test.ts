import { expect, test, setDefaultTimeout } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";
import { EditorState } from "prosemirror-state";
import { freshCommentId, findStoryBody } from "@stll/docx-core/ops";
import { validateDocxPackage } from "@stll/docx-core";
import { serializeCanonicalSave } from "../docx/canonicalSave";
import { FOLIO_DOCX_SERIALIZATION_MODE } from "../types/docxSerialization";
import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { canonicalCommentBody, compileCanonicalComments } from "./canonicalComments";
import { createCanonicalSession, publishCanonicalProjection } from "./canonicalSession";
import type { Document, Paragraph } from "../types/document";
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

const COMMENTS_SOURCE_PATHS = [
  "word/comments.xml",
  "word/Comments.xml",
  "word/COMMENTS.XML",
] as const;

const openCommentSource = async (path: (typeof COMMENTS_SOURCE_PATHS)[number]) => {
  const source = setup();
  const id = freshCommentId(source.session.document).unwrap();
  const block = source.session.document.package.document.content.at(0);
  if (block?.type !== "paragraph" || !block.paraId) throw new Error("Missing source paragraph");
  source.apply({
    type: "create",
    comment: {
      id,
      author: "source",
      content: canonicalCommentBody(source.session.document, "source"),
    },
    anchor: { kind: "point", at: { story: "main", blockId: block.paraId, offset: 1 } },
  });
  const zip = await JSZip.loadAsync(await createDocx(source.session.document));
  const comments = await zip.file("word/comments.xml")?.async("text");
  if (!comments) throw new Error("Missing source comments part");
  expect(zip.file("word/commentsExtended.xml")).toBeNull();
  zip.remove("word/comments.xml");
  zip.file(path, comments);
  const contentTypes = await zip.file("[Content_Types].xml")?.async("text");
  const relationships = await zip.file("word/_rels/document.xml.rels")?.async("text");
  if (!contentTypes || !relationships) throw new Error("Missing source comment packaging");
  zip.file("[Content_Types].xml", contentTypes.replace("/word/comments.xml", `/${path}`));
  zip.file(
    "word/_rels/document.xml.rels",
    relationships.replace('Target="comments.xml"', `Target="${path.slice("word/".length)}"`),
  );
  const buffer = await zip.generateAsync({ type: "arraybuffer" });
  expect(await validateDocxPackage(new Uint8Array(buffer))).toEqual({ valid: true });
  return { document: await parseDocx(buffer, { preloadFonts: false }), id };
};

const assertCanonicalCommentSave = async (session: ReturnType<typeof setup>["session"]) => {
  const snapshot = session.captureSaveSnapshot();
  expect(snapshot.structure).toBe("stable");
  for (const mode of Object.values(FOLIO_DOCX_SERIALIZATION_MODE)) {
    const saved = await serializeCanonicalSave({
      snapshot,
      featureFlags: { selectiveSave: true },
      options: { mode },
    });
    expect(await validateDocxPackage(new Uint8Array(saved.buffer))).toEqual({ valid: true });
    const reopened = await parseDocx(saved.buffer, { preloadFonts: false });
    // The package oracle includes root/reply status and body range/reference
    // anchors, rather than only the visible comment text.
    expect(
      describePackageDifferences(snapshot.document, reopened),
      `Canonical comment save mode: ${mode}`,
    ).toEqual({
      messages: [],
      omitted: 0,
    });
  }
};

test.each(["header", "footnote"] as const)(
  "selective comment save preserves created and deleted %s anchors",
  async (storyKind) => {
    const document: Document = seed();
    const block = {
      type: "paragraph",
      paraId: "ABCDEF02",
      content: [{ type: "run", content: [{ type: "text", text: "Secondary story" }] }],
    } satisfies Paragraph;
    const story =
      storyKind === "header"
        ? ({ kind: "header", rId: "rIdHeader" } as const)
        : ({ kind: "footnote", id: 1 } as const);
    if (storyKind === "header") {
      document.package.headers = new Map([
        ["rIdHeader", { type: "header", hdrFtrType: "default", content: [block] }],
      ]);
      document.package.document.finalSectionProperties = {
        headerReferences: [{ type: "default", rId: "rIdHeader" }],
      };
    } else {
      document.package.footnotes = [{ type: "footnote", id: 1, content: [block] }];
      const body = document.package.document.content.at(0);
      if (body?.type !== "paragraph") throw new Error("Missing body paragraph");
      body.content.push({ type: "run", content: [{ type: "footnoteRef", id: 1 }] });
    }
    const source = await parseDocx(await createDocx(document), { preloadFonts: false });
    const editor = setup(source);
    const id = freshCommentId(editor.session.document).unwrap();
    const target = findStoryBody(editor.session.document, story)?.content.at(0);
    if (target?.type !== "paragraph" || !target.paraId) throw new Error("Missing story paragraph");
    editor.apply({
      type: "create",
      comment: {
        id,
        author: "Reviewer",
        content: canonicalCommentBody(editor.session.document, "Secondary comment"),
      },
      anchor: { kind: "point", at: { story, blockId: target.paraId, offset: 2 } },
    });
    const saveAndReopen = async (current: ReturnType<typeof setup>) => {
      const snapshot = current.session.captureSaveSnapshot();
      const saved = await serializeCanonicalSave({
        snapshot,
        featureFlags: { selectiveSave: true },
        options: { mode: FOLIO_DOCX_SERIALIZATION_MODE.preferSelective },
      });
      expect(snapshot.changedBlockIds).toContain(target.paraId);
      expect(saved.diagnostics).toEqual(
        storyKind === "header" ? [{ type: "selectiveSaveRefused", part: "word/document.xml" }] : [],
      );
      const reopened = await parseDocx(saved.buffer, { preloadFonts: false });
      expect(findStoryBody(reopened, story)?.content).toEqual(
        findStoryBody(snapshot.document, story)?.content,
      );
      expect(describePackageDifferences(snapshot.document, reopened)).toEqual({
        messages: [],
        omitted: 0,
      });
      return reopened;
    };
    const created = await saveAndReopen(editor);
    const deletion = setup(created);
    deletion.apply({ type: "delete", id });
    await saveAndReopen(deletion);
    deletion.history("undo");
    await saveAndReopen(deletion);
    deletion.history("redo");
    await saveAndReopen(deletion);
  },
);

test("generated point/range comment histories preserve exact undo, projection and canonical snapshot save", async () => {
  await assertProperty(
    fc.asyncProperty(
      fc.array(fc.stringMatching(/^[a-z]{1,12}$/u), { minLength: 1, maxLength: 6 }),
      fc.constantFrom("point", "range"),
      fc.constantFrom(...COMMENTS_SOURCE_PATHS),
      async (texts, kind, sourcePath) => {
        const source = await openCommentSource(sourcePath);
        const editor = setup(source.document);
        const original = editor.session.document;
        editor.apply({ type: "delete", id: source.id });
        expect(editor.session.captureSaveSnapshot().structure).toBe("stable");
        const deletion = await serializeCanonicalSave({
          snapshot: editor.session.captureSaveSnapshot(),
          featureFlags: { selectiveSave: true },
          options: { mode: FOLIO_DOCX_SERIALIZATION_MODE.preferSelective },
        });
        expect(deletion.diagnostics).toEqual([]);
        expect(await validateDocxPackage(new Uint8Array(deletion.buffer))).toEqual({ valid: true });
        const deletedZip = await JSZip.loadAsync(deletion.buffer);
        expect(
          Object.keys(deletedZip.files).filter(
            (path) => path.toLowerCase() === "word/comments.xml",
          ),
        ).toEqual([sourcePath]);
        expect(
          describePackageDifferences(
            editor.session.document,
            await parseDocx(deletion.buffer, { preloadFonts: false }),
          ),
        ).toEqual({ messages: [], omitted: 0 });
        await assertCanonicalCommentSave(editor.session);
        let steps = 1;
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
          await assertCanonicalCommentSave(editor.session);
          editor.apply({ type: "delete", id });
          expect(
            editor.session.document.package.document.comments?.some(
              (comment) => comment.id === id || comment.id === replyId,
            ) ?? false,
          ).toBe(false);
          editor.history("undo");
          expect(editor.session.document).toEqual(committed);
          await assertCanonicalCommentSave(editor.session);
          editor.history("redo");
          await assertCanonicalCommentSave(editor.session);
          steps += 1;
        }
        const final = editor.session.document;
        for (let index = 0; index < steps; index += 1) editor.history("undo");
        expect(editor.session.document).toEqual(original);
        await assertCanonicalCommentSave(editor.session);
        for (let index = 0; index < steps; index += 1) editor.history("redo");
        expect(editor.session.document).toEqual(final);
        await assertCanonicalCommentSave(editor.session);
        expect(editor.session.projection.doc.textContent).toBe("A😀éB");
      },
    ),
    {
      numRuns: 12,
      seed: 20261005,
      examples: COMMENTS_SOURCE_PATHS.map((path) => [["a"], "point" as const, path]),
    },
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
  await assertCanonicalCommentSave(editor.session);
});

test("a loaded root comment without source anchors activates, stays read-only and survives save", async () => {
  const source = seed();
  source.package.document.comments = [
    {
      id: 7,
      author: "Reviewer",
      content: [
        {
          type: "paragraph",
          paraId: "ABCDEF02",
          content: [{ type: "run", content: [{ type: "text", text: "Orphan" }] }],
        },
      ],
    },
  ];
  const loaded = await parseDocx(await createDocx(source));
  expect(loaded.package.document.comments?.map(({ id }) => id)).toEqual([7]);
  const editor = setup(loaded);
  const refusedOnOrphan = (command: Parameters<typeof compileCanonicalComments>[0]["command"]) => {
    const compiled = compileCanonicalComments({ document: editor.session.document, command });
    if (compiled.isErr()) return true;
    const state = EditorState.create({ doc: editor.session.projection.doc });
    return editor.session.prepareOperations(state, compiled.value.ops).isErr();
  };
  expect(refusedOnOrphan({ type: "resolve", id: 7, status: "resolved" })).toBe(true);
  expect(refusedOnOrphan({ type: "delete", id: 7 })).toBe(true);
  expect(
    refusedOnOrphan({
      type: "create",
      anchor: { kind: "reply", parentId: 7 },
      comment: {
        id: freshCommentId(editor.session.document).unwrap(),
        author: "Reviewer",
        done: false,
        content: canonicalCommentBody(editor.session.document, "reply"),
      },
    }),
  ).toBe(true);
  // Other comments stay fully editable beside the preserved one.
  editor.create("anchored", "range");
  expect(editor.session.document.package.document.comments?.map(({ id }) => id)).toContain(7);
  await assertCanonicalCommentSave(editor.session);
});
