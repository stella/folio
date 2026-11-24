/**
 * A comment added on text another comment's range runs past puts its
 * reference inside that range: a range is one stretch, so a save writes the
 * reference between the outer range's start and end, and reopening reads it
 * there. The live document holds it there too, so removing the text after the
 * reference leaves the outer range ending after it, saved or not.
 */

import { describe, expect, test } from "bun:test";

import { MAX_REVISION_ID } from "@stll/docx-core/model";

import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
  type FolioDocumentOperationMode,
} from "../document-operations";
import { ensureParaIds } from "../docx/ensureParaIds";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import type { Document } from "../types/document";
import { FolioDocxReviewer } from "./headless";
import { hashFolioAIBlockText } from "./snapshot";

const MODES = ["direct", "tracked-changes"] as const;
const SESSIONS = ["canonical"] as const;
type Session = (typeof SESSIONS)[number];

const open = async (bytes: ArrayBuffer | Uint8Array, session: Session) =>
  FolioDocxReviewer.fromBuffer(bytes, { author: "Test", session });

const apply = (
  reviewer: FolioDocxReviewer,
  mode: FolioDocumentOperationMode,
  operations: FolioDocumentOperation[],
) => {
  const result = reviewer.applyDocumentOperations(
    { version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION, mode, operations },
    { undefinedReferences: "refuse" },
  );
  expect(result.skipped).toEqual([]);
};

const commentText = (document: Document, id: number): string =>
  (document.package.document.comments ?? [])
    .find((comment) => comment.id === id)
    ?.content.flatMap((paragraph) => paragraph.content)
    .flatMap((child) => (child.type === "run" ? child.content : []))
    .map((item) => (item.type === "text" ? item.text : ""))
    .join("") ?? String(id);

/** The commented paragraph's words and comment markers, named by comment text. */
const markers = (document: Document) =>
  document.package.document.content.slice(0, 1).flatMap((block) =>
    block.type === "paragraph"
      ? block.content.flatMap((child) => {
          if (child.type === "run")
            return child.content.flatMap((item) =>
              item.type === "text" ? [item.text] : [],
            );
          if (
            child.type === "commentRangeStart" ||
            child.type === "commentRangeEnd" ||
            child.type === "commentReference"
          )
            return [`${child.type} ${commentText(document, child.id)}`];
          return [];
        })
      : [],
  );

/** "Outer" covers the paragraph; "Inner" covers its first word only. */
const nested = async (session: Session, reservedId?: number) => {
  const document = fromMarkdown(
    reservedId === undefined
      ? "Clause one tail."
      : "Clause one tail.\n\nOther paragraph.",
  );
  if (reservedId !== undefined) {
    const other = document.package.document.content.at(1);
    if (other?.type !== "paragraph") throw new Error("Missing allocator seed");
    other.content = [
      { type: "commentRangeStart", id: reservedId },
      ...other.content,
      { type: "commentRangeEnd", id: reservedId },
      { type: "run", content: [{ type: "commentReference", id: reservedId }] },
    ];
    document.package.document.comments = [
      {
        id: reservedId,
        author: "Test",
        date: "2026-01-01T00:00:00Z",
        anchorKind: "range",
        content: [
          {
            type: "paragraph",
            content: [
              {
                type: "run",
                content: [{ type: "text", text: "Already occupied" }],
              },
            ],
          },
        ],
      },
    ];
  }
  const reviewer = await open(
    (await ensureParaIds(await createDocx(document))).docx,
    session,
  );
  const block = reviewer.snapshot().blocks[0];
  if (!block) throw new Error("Missing fixture");
  apply(reviewer, "direct", [
    {
      id: "outer",
      type: "commentOnBlock",
      blockId: block.id,
      comment: { text: "Outer" },
    },
  ]);
  apply(reviewer, "direct", [
    {
      id: "inner",
      type: "commentOnRange",
      range: {
        type: "textRange",
        story: "main",
        blockId: block.id,
        startOffset: 0,
        endOffset: 6,
        selectedTextHash: hashFolioAIBlockText("Clause"),
      },
      comment: { text: "Inner" },
    },
  ]);
  return { reviewer, blockId: block.id };
};

const removeTail = (
  reviewer: FolioDocxReviewer,
  mode: FolioDocumentOperationMode,
  blockId: string,
) => {
  apply(reviewer, mode, [
    {
      id: "tail",
      type: "replaceInBlock",
      blockId,
      find: " one tail.",
      replace: "",
    },
  ]);
  return markers(
    mode === "direct"
      ? reviewer.toDocument()
      : reviewer.toResolvedDocument("accept"),
  );
};

const commentIdsByText = (reviewer: FolioDocxReviewer) => {
  const document = reviewer.toDocument();
  return Object.fromEntries(
    (document.package.document.comments ?? []).map(({ id }) => [
      commentText(document, id),
      id,
    ]),
  );
};

// The legacy allocator starts at 1; the canonical one is one past every
// annotation id, so a package without any starts at 0. Both wrap to the
// lowest free id once the highest is taken.
const FIRST_ID = { legacy: 1, canonical: 0 } as const satisfies Record<
  Session,
  number
>;

// Range starts at one position carry no order: the legacy session writes
// them by id, the canonical session puts the newer comment's first.
const startsOf = (session: Session, byId: readonly string[]) =>
  session === "legacy" ? byId : ["Inner", "Outer"];

const ID_ORDERS = [
  {
    name: "ascending IDs",
    reservedId: undefined,
    outerId: (session: Session) => FIRST_ID[session],
    innerId: (session: Session) => FIRST_ID[session] + 1,
    starts: ["Outer", "Inner"],
  },
  {
    name: "wrapped IDs",
    reservedId: MAX_REVISION_ID - 1,
    outerId: () => MAX_REVISION_ID,
    innerId: (session: Session) => FIRST_ID[session],
    starts: ["Inner", "Outer"],
  },
] as const;

describe("a comment reference put inside another comment's range", () => {
  for (const session of SESSIONS)
    for (const mode of MODES) {
      for (const { name, reservedId, outerId, innerId, starts } of ID_ORDERS) {
        test(`stays inside it after removing the tail (${session}, ${mode}, ${name})`, async () => {
          const { reviewer, blockId } = await nested(session, reservedId);
          const commentIds =
            reservedId === undefined
              ? { Outer: outerId(session), Inner: innerId(session) }
              : {
                  Outer: outerId(session),
                  Inner: innerId(session),
                  "Already occupied": reservedId,
                };
          expect(commentIdsByText(reviewer)).toEqual(commentIds);
          const liveBeforeSave = markers(reviewer.toDocument());
          const reopened = await open(await reviewer.toBuffer(), session);
          expect(markers(reviewer.toDocument())).toEqual(liveBeforeSave);
          expect(commentIdsByText(reopened)).toEqual(commentIds);
          const expected = [
            ...startsOf(session, starts).map(
              (comment) => `commentRangeStart ${comment}`,
            ),
            "Clause",
            "commentRangeEnd Inner",
            "commentReference Inner",
            "commentRangeEnd Outer",
            "commentReference Outer",
          ];
          expect(removeTail(reviewer, mode, blockId)).toEqual(expected);
          expect(removeTail(reopened, mode, blockId)).toEqual(expected);
        });
      }
    }
});

/** The ids of the tracked insertions and deletions anywhere in `value`. */
const trackedIds = (value: unknown, ids = new Set<number>()): Set<number> => {
  if (Array.isArray(value)) {
    for (const item of value) trackedIds(item, ids);
    return ids;
  }
  if (typeof value !== "object" || value === null) return ids;
  const type: unknown = Reflect.get(value, "type");
  const info: unknown = Reflect.get(value, "info");
  if (
    (type === "insertion" || type === "deletion") &&
    typeof info === "object" &&
    info !== null
  ) {
    const id: unknown = Reflect.get(info, "id");
    if (typeof id === "number") ids.add(id);
  }
  for (const child of Object.values(value)) trackedIds(child, ids);
  return ids;
};

test.each(SESSIONS)(
  "past the highest id, a comment and a tracked edit of one batch take different ids (%s)",
  async (session) => {
    const { reviewer } = await nested(session, MAX_REVISION_ID - 1);
    const other = reviewer.snapshot().blocks[1];
    if (!other) throw new Error("Missing second paragraph");
    apply(reviewer, "tracked-changes", [
      {
        id: "comment",
        type: "commentOnBlock",
        blockId: other.id,
        comment: { text: "New" },
      },
      {
        id: "edit",
        type: "replaceInBlock",
        blockId: other.id,
        find: "Other",
        replace: "Another",
      },
    ]);
    const document = reviewer.toDocument();
    const comments = new Set(
      (document.package.document.comments ?? []).map(({ id }) => id),
    );
    const revisions = trackedIds(document.package.document.content);
    expect(revisions.size).toBeGreaterThan(0);
    expect([...revisions].filter((id) => comments.has(id))).toEqual([]);
  },
);

describe("a block comment over two comments on the same words", () => {
  test.each(
    SESSIONS.flatMap((session) =>
      MODES.map((mode) => [session, mode] as const),
    ),
  )(
    "keeps both references inside its range (%s, %s)",
    async (session, mode) => {
      const reviewer = await open(
        (
          await ensureParaIds(
            await createDocx(fromMarkdown("Clause one tail.")),
          )
        ).docx,
        session,
      );
      const block = reviewer.snapshot().blocks[0];
      if (!block) throw new Error("Missing fixture");
      for (const text of ["First", "Second"])
        apply(reviewer, mode, [
          {
            id: text,
            type: "commentOnRange",
            range: {
              type: "textRange",
              story: "main",
              blockId: block.id,
              startOffset: 0,
              endOffset: 6,
              selectedTextHash: hashFolioAIBlockText("Clause"),
            },
            comment: { text },
          },
        ]);
      apply(reviewer, mode, [
        {
          id: "outer",
          type: "commentOnBlock",
          blockId: block.id,
          comment: { text: "Outer" },
        },
      ]);
      const references = markers(reviewer.toDocument()).filter((marker) =>
        marker.startsWith("commentReference"),
      );
      expect(references.toSorted()).toEqual([
        "commentReference First",
        "commentReference Outer",
        "commentReference Second",
      ]);
    },
  );
});
