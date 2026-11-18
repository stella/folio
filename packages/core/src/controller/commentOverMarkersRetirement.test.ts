/**
 * A comment whose covered text is all deleted retires, also when bookmarks or
 * other range boundaries sat inside its range: they mark positions, not
 * content. Both engines, after a save.
 */
import { afterEach, expect, setSystemTime, test } from "bun:test";
import { EditorState } from "prosemirror-state";

import { ZipArchive } from "@stll/docx-core/zip";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { getTrackedChangeStatsFromDoc } from "../ai-edits/read";
import { createFolioAIEditSnapshot } from "../ai-edits/snapshot";
import { CanonicalPublicOperations } from "../canonical/canonicalPublicOperations";
import {
  type CanonicalCommit,
  createCanonicalSession,
  publishCanonicalProjection,
} from "../canonical/canonicalSession";
import { serializeCanonicalSave } from "../docx/canonicalSave";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import {
  type CommentMarkerSlot,
  visitCommentMarkers,
} from "../docx/commentAnchorIndex";
import { parseDocx } from "../docx/parser";
import { RELATIONSHIP_TYPES } from "../docx/relsParser";
import { expectBookmarkBoundaryAttrs } from "../prosemirror/bookmarkBoundaryAttrs";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { readModel } from "../types/documentModel";

afterEach(() => setSystemTime());

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const COMMENTS_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";

const run = (text: string) =>
  `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

const fixture = async (): Promise<ArrayBuffer> => {
  const zip = new ZipArchive();
  zip.set(
    "[Content_Types].xml",
    `${HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
      `<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>` +
      `</Types>`,
  );
  zip.set(
    "_rels/.rels",
    `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="${RELATIONSHIP_TYPES.officeDocument}" Target="word/document.xml"/></Relationships>`,
  );
  zip.set(
    "word/_rels/document.xml.rels",
    `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId9" Type="${COMMENTS_TYPE}" Target="comments.xml"/></Relationships>`,
  );
  zip.set(
    "word/document.xml",
    `${HEAD}<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>` +
      `<w:p w14:paraId="1A2B3C4D"><w:commentRangeStart w:id="0"/><w:bookmarkStart w:id="5" w:name="_GoBack"/>` +
      `${run("This")}<w:bookmarkEnd w:id="5"/><w:commentRangeEnd w:id="0"/>` +
      `<w:r><w:commentReference w:id="0"/></w:r>${run(" stays.")}</w:p>` +
      `<w:p w14:paraId="2B3C4D5E">${run("Tail.")}</w:p>` +
      `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`,
  );
  zip.set(
    "word/comments.xml",
    `${HEAD}<w:comments xmlns:w="${W}"><w:comment w:id="0" w:author="Reviewer" w:date="2024-01-01T00:00:00Z" w:initials="R">` +
      `<w:p>${run("Check this.")}</w:p></w:comment></w:comments>`,
  );
  return (await zip.generate({ compression: "store" })).buffer;
};

const DELETE_COMMENTED_TEXT = {
  id: "drop",
  type: "replaceInBlock",
  blockId: "1A2B3C4D",
  find: "This",
  replace: "",
} as const;

test("the default engine drops the comment from the save", async () => {
  const reviewer = await FolioDocxReviewer.fromBuffer(await fixture());
  expect(reviewer.getComments()).toHaveLength(1);
  const result = reviewer.applyOperations([DELETE_COMMENTED_TEXT], {
    undefinedReferences: "refuse",
    mode: "direct",
  });
  expect(result.skipped).toEqual([]);
  const reopened = await FolioDocxReviewer.fromBuffer(
    await reviewer.toBuffer(),
  );
  expect(reopened.getComments()).toEqual([]);
  expect(reopened.getContent().map(({ text }) => text)).toContain(" stays.");
});

test("the canonical session drops the comment from the save", async () => {
  const prepared = (await prepareCanonicalDocxInput(await fixture())).unwrap();
  const document = await parseDocx(prepared, {
    preloadFonts: false,
    detectVariables: false,
  });
  const session = createCanonicalSession(
    document,
    document.package.styles,
  ).unwrap();
  let state = EditorState.create({ doc: session.projection.doc });
  const publish = (commit: CanonicalCommit) => {
    const published = publishCanonicalProjection({ state, commit, session });
    if (published.isErr()) return false;
    state = published.value.state;
    return true;
  };
  const executor = new CanonicalPublicOperations({
    session,
    getState: () => state,
    publish,
  });
  const result = executor.apply({
    snapshot: createFolioAIEditSnapshot(state.doc),
    batch: { version: 1, mode: "direct", operations: [DELETE_COMMENTED_TEXT] },
    undefinedReferences: "refuse",
  });
  expect(result.applied).toHaveLength(1);
  const saved = await serializeCanonicalSave({
    snapshot: session.captureSaveSnapshot(),
  });
  const reopened = await FolioDocxReviewer.fromBuffer(saved.buffer);
  expect(reopened.getComments()).toEqual([]);
  expect(reopened.getContent().map(({ text }) => text)).toContain(" stays.");
});

test("a source comment whose range holds only a bookmark stays a point comment", async () => {
  const zip = await ZipArchive.read(await fixture());
  const xml = (await zip.entry("word/document.xml")?.text()) ?? "";
  zip.set("word/document.xml", xml.replace(run("This"), ""));
  const reviewer = await FolioDocxReviewer.fromBuffer(
    (await zip.generate({ compression: "store" })).buffer,
  );
  const reopened = await FolioDocxReviewer.fromBuffer(
    await reviewer.toBuffer(),
  );
  expect(reopened.getComments()).toHaveLength(1);
});

type RetirementCase =
  | "rangedWithTail"
  | "rangeElsewhere"
  | "wholeParagraph"
  | "pointOnDeletedWord"
  | "bookmarkOnlyPoint";
const retirementFixture = async (shape: RetirementCase) => {
  const zip = await ZipArchive.read(await fixture());
  const source = await zip.entry("word/document.xml")?.text();
  if (source === undefined) throw new Error("Missing fixture document XML");
  const end =
    '<w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r>';
  const xml = (() => {
    switch (shape) {
      case "rangedWithTail":
        return source;
      case "rangeElsewhere":
        return source
          .replace(run(" stays."), "")
          .replace(end, "")
          .replace(run("Tail."), run("Tail.") + end);
      case "wholeParagraph":
        return source.replace(run(" stays."), "");
      case "pointOnDeletedWord":
        return source
          .replace('<w:commentRangeStart w:id="0"/>', "")
          .replace('<w:commentRangeEnd w:id="0"/>', "");
      case "bookmarkOnlyPoint":
        return source.replace(run("This"), "");
      default: {
        const exhaustive: never = shape;
        throw new Error(String(exhaustive));
      }
    }
  })();
  zip.set("word/document.xml", xml);
  return (await zip.generate({ compression: "store" })).buffer;
};

const savedCommentFacts = async (bytes: ArrayBuffer) => {
  const parsed = await parseDocx(bytes, {
    preloadFonts: false,
    detectVariables: false,
  });
  const reviewer = await FolioDocxReviewer.fromBuffer(bytes.slice(0));
  const markers: CommentMarkerSlot["item"][] = [];
  visitCommentMarkers(parsed.package.document.content, ({ item }) =>
    markers.push(item),
  );
  const bookmarks: { type: "start" | "end"; id: number; name?: string }[] = [];
  toProseDoc(parsed).descendants((node) => {
    if (node.type.name === "bookmarkBoundary") {
      const { type, id, name } = expectBookmarkBoundaryAttrs(node);
      bookmarks.push({ type, id, ...(name === undefined ? {} : { name }) });
    }
  });
  return {
    comments: reviewer.getComments(),
    definitions: parsed.package.document.comments ?? [],
    markers,
    bookmarks,
    text: reviewer.getContentAsText(),
  };
};

for (const shape of [
  "rangedWithTail",
  "rangeElsewhere",
  "wholeParagraph",
  "pointOnDeletedWord",
  "bookmarkOnlyPoint",
] as const)
  for (const mode of ["direct", "tracked-changes", "suggested"] as const) {
    test(`${shape}, ${mode}: pure deletion follows parsed comment ownership or refuses unchanged`, async () => {
      setSystemTime(new Date("2026-01-02T03:04:05Z"));
      const prepared = (
        await prepareCanonicalDocxInput(await retirementFixture(shape))
      ).unwrap();
      const source = await parseDocx(prepared, {
        preloadFonts: false,
        detectVariables: false,
      });
      const session = createCanonicalSession(
        source,
        source.package.styles,
      ).unwrap();
      let state = EditorState.create({ doc: session.projection.doc });
      const publish = (commit: CanonicalCommit) => {
        state = publishCanonicalProjection({ state, commit, session }).unwrap()
          .state;
        return true;
      };
      const executor = new CanonicalPublicOperations({
        session,
        getState: () => state,
        publish,
      });
      const operation = {
        ...DELETE_COMMENTED_TEXT,
        find: shape === "bookmarkOnlyPoint" ? " stays." : "This",
      };
      const apply = () =>
        executor.apply({
          snapshot: createFolioAIEditSnapshot(state.doc),
          batch: { version: 1, mode, operations: [operation] },
          undefinedReferences: "refuse",
        });
      const before = session.document;
      const snapshot = session.captureSaveSnapshot();
      const saved = await serializeCanonicalSave({ snapshot });
      const initial = await savedCommentFacts(saved.buffer);
      expect(initial.comments).toHaveLength(1);
      expect(initial.bookmarks).toHaveLength(2);
      const result = apply();
      if (shape === "wholeParagraph" || shape === "pointOnDeletedWord") {
        expect(result.applied).toEqual([]);
        expect(result.undoHandle).toBeNull();
        expect(result.skipped).toMatchObject([
          {
            reason: "unsupportedBlock",
            canonicalRefusal: { gap: "comment-replacement-alignment" },
          },
        ]);
        expect(session.document).toBe(before);
        expect(session.version).toBe(0);
        expect(session.canUndo).toBe(false);
        expect(session.canRedo).toBe(false);
        expect(session.captureSaveSnapshot()).toEqual(snapshot);
        expect(
          new Uint8Array(
            (
              await serializeCanonicalSave({
                snapshot: session.captureSaveSnapshot(),
              })
            ).buffer,
          ),
        ).toEqual(new Uint8Array(saved.buffer));
        return;
      }
      expect(result.applied).toHaveLength(1);
      expect(result.skipped).toEqual([]);
      if (result.undoHandle === null)
        throw new Error("Deletion returned no undo handle");
      expect(executor.undo(result.undoHandle)).toMatchObject({
        status: "undone",
      });
      expect(session.document).toEqual(before);
      const reapplied = apply();
      expect(reapplied.applied).toHaveLength(1);
      expect(reapplied.skipped).toEqual([]);
      if (mode !== "direct") {
        const rejected = createCanonicalSession(
          readModel(session.document),
        ).unwrap();
        const rejectedState = EditorState.create({
          doc: rejected.projection.doc,
        });
        const suggestionId = reapplied.applied.at(0)?.suggestionId;
        if (mode === "suggested" && suggestionId === undefined)
          throw new Error("Deletion returned no suggestion id");
        const rejectedCommit =
          mode === "suggested" && suggestionId !== undefined
            ? rejected
                .prepareSuggestion(rejectedState, {
                  suggestionId,
                  resolution: "reject",
                  author: "Reviewer",
                  date: "2026-01-02T03:04:05Z",
                })
                .unwrap()
            : rejected
                .prepareResolve(rejectedState, {
                  revisionIds: getTrackedChangeStatsFromDoc(rejectedState.doc)
                    .ids,
                  resolution: "reject",
                })
                .unwrap();
        publishCanonicalProjection({
          session: rejected,
          state: rejectedState,
          commit: rejectedCommit,
        }).unwrap();
        expect(
          await savedCommentFacts(
            (
              await serializeCanonicalSave({
                snapshot: rejected.captureSaveSnapshot(),
              })
            ).buffer,
          ),
        ).toEqual(initial);
        if (mode === "suggested" && suggestionId !== undefined)
          publish(
            session
              .prepareSuggestion(state, {
                suggestionId,
                resolution: "accept",
                author: "Reviewer",
                date: "2026-01-02T03:04:05Z",
              })
              .unwrap(),
          );
        publish(
          session
            .prepareResolve(state, {
              revisionIds: getTrackedChangeStatsFromDoc(state.doc).ids,
              resolution: "accept",
            })
            .unwrap(),
        );
      }
      const final = await savedCommentFacts(
        (
          await serializeCanonicalSave({
            snapshot: session.captureSaveSnapshot(),
          })
        ).buffer,
      );
      expect(final.bookmarks).toEqual(initial.bookmarks);
      if (shape === "rangedWithTail") {
        expect(final.comments).toEqual([]);
        expect(final.definitions).toEqual([]);
        expect(final.markers).toEqual([]);
      } else {
        expect(final.definitions).toEqual(initial.definitions);
        expect(final.markers).toEqual(initial.markers);
        expect(final.comments.map(({ anchoredText }) => anchoredText)).toEqual([
          shape === "rangeElsewhere" ? "Tail." : "",
        ]);
      }
    });
  }
