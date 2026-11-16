import { afterEach, describe, expect, setSystemTime, test } from "bun:test";

import { FolioDocxReviewer } from "../ai-edits/headless";
import {
  canonicalCommentBody,
  compileCanonicalComments,
} from "../canonical/canonicalComments";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";

const SOURCE = "The Supplier delivers the goods on time and in good order.";
const REPLACEMENT =
  "Invoice delivery payment agreement supplier buyer payment term.";
afterEach(() => setSystemTime());
const COMMENT_CASES = {
  prefix: { quotes: ["The"], type: "alignment" },
  interior: { quotes: ["time"], type: "alignment" },
  suffix: { quotes: ["order."], type: "supported" },
  whole: { quotes: [SOURCE], type: "supported" },
  disjoint: { quotes: ["The", "time"], type: "alignment" },
  nested: { quotes: [SOURCE, "time"], type: "alignment" },
  overlapping: { quotes: ["on time", "time and"], type: "alignment" },
} as const;

const open = (buffer: ArrayBuffer, session: HeadlessReviewerSession) =>
  FolioDocxReviewer.fromBuffer(buffer.slice(0), {
    session,
    author: "Reviewer",
  });

const apply = (
  reviewer: FolioDocxReviewer,
  operation: FolioDocumentOperation,
  mode: "direct" | "tracked-changes",
) => {
  const result = reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode,
      operations: [operation],
    },
    { undefinedReferences: "refuse" },
  );
  expect(result.skipped).toEqual([]);
  expect(result.applied).toHaveLength(1);
  return result;
};

const comments = (reviewer: FolioDocxReviewer) =>
  reviewer.getComments().map(({ id, text, anchoredText, replies }) => ({
    id,
    text,
    anchoredText,
    replies: replies.map(({ id: replyId, text: replyText }) => ({
      id: replyId,
      text: replyText,
    })),
  }));

describe("canonical replacement comments from the public producer", () => {
  for (const offset of [0, 4, SOURCE.length]) {
    test(`point at ${offset}: replacement retains its carriers or refuses unchanged`, async () => {
      setSystemTime(new Date("2026-01-02T03:04:05Z"));
      const model = fromMarkdown(SOURCE);
      const paragraph = model.package.document.content.at(0);
      if (paragraph?.type !== "paragraph")
        throw new Error("Source paragraph is missing");
      paragraph.paraId = "12345678";
      const compiled = compileCanonicalComments({
        document: model,
        command: {
          type: "create",
          comment: {
            id: 0,
            author: "Reviewer",
            content: canonicalCommentBody(model, "Point"),
          },
          anchor: {
            kind: "point",
            at: { story: "main", blockId: paragraph.paraId, offset },
          },
        },
      });
      if (compiled.isErr()) throw compiled.error;
      const prepared = await prepareCanonicalDocxInput(
        await createDocx(compiled.value.document),
      );
      if (prepared.isErr()) throw prepared.error;
      for (const text of [REPLACEMENT, ""]) {
        for (const mode of ["direct", "tracked-changes"] as const) {
          const reviewer = await open(
            prepared.value.slice().buffer,
            "canonical",
          );
          const original = comments(reviewer);
          expect(original).toHaveLength(1);
          const before = reviewer.toDocument();
          const saved = await reviewer.toBuffer();
          const result = reviewer.applyDocumentOperations(
            {
              version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
              mode,
              operations: [
                {
                  id: "replace",
                  type: "replaceBlock",
                  blockId: paragraph.paraId,
                  text,
                },
              ],
            },
            { undefinedReferences: "refuse" },
          );
          if (result.applied.length === 0) {
            expect(result.skipped).toMatchObject([
              {
                reason: "unsupportedBlock",
                canonicalRefusal: { gap: "comment-replacement-alignment" },
              },
            ]);
            expect(reviewer.toDocument()).toEqual(before);
            expect(new Uint8Array(await reviewer.toBuffer())).toEqual(
              new Uint8Array(saved),
            );
            continue;
          }
          expect(result.skipped).toEqual([]);
          const pending = await reviewer.toBuffer();
          expect(reviewer.acceptAll().skipped).toEqual([]);
          expect(comments(reviewer)).toEqual(original);
          expect(
            comments(await open(await reviewer.toBuffer(), "canonical")),
          ).toEqual(original);
          if (mode === "tracked-changes") {
            const rejected = await open(pending, "canonical");
            expect(rejected.rejectAll().skipped).toEqual([]);
            expect(comments(rejected)).toEqual(original);
          }
        }
      }
    });
  }
  for (const [name, scenario] of Object.entries(COMMENT_CASES)) {
    test(`${name}: supported replacements preserve threads; unsupported alignment refuses unchanged`, async () => {
      // Save updates package modification dates, independently of any edit.
      setSystemTime(new Date("2026-01-02T03:04:05Z"));
      const prepared = await prepareCanonicalDocxInput(
        await createDocx(fromMarkdown(SOURCE)),
      );
      if (prepared.isErr()) throw prepared.error;
      // The legacy public producer creates source threads, including replies;
      // every editor opens the same source package and existing comment IDs.
      const producer = await open(prepared.value.slice().buffer, "legacy");
      const block = producer.getContent().at(0);
      if (!block) throw new Error("Source paragraph is missing");
      for (const [index, quote] of scenario.quotes.entries()) {
        apply(
          producer,
          {
            id: `comment-${index}`,
            type: "commentOnBlock",
            blockId: block.id,
            quote,
            comment: { text: `Note ${index}` },
          },
          "direct",
        );
        const comment = producer.getComments().at(-1);
        if (!comment) throw new Error("Produced comment is missing");
        expect(
          producer.replyTo(comment, { text: `Reply ${index}` }),
        ).not.toBeNull();
      }
      const bytes = await producer.toBuffer();
      const operation = {
        id: "replace",
        type: "replaceBlock",
        blockId: block.id,
        text: REPLACEMENT,
      } as const;
      const direct = await open(bytes, "canonical");
      const tracked = await open(bytes, "canonical");
      const legacy = await open(bytes, "legacy");
      const original = comments(tracked);
      expect(original).toHaveLength(scenario.quotes.length);
      apply(legacy, operation, "direct");
      if (scenario.type === "alignment") {
        for (const mode of ["direct", "tracked-changes"] as const) {
          const reviewer = await open(bytes, "canonical");
          const before = reviewer.toDocument();
          const saved = await reviewer.toBuffer();
          const result = reviewer.applyDocumentOperations(
            {
              version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
              mode,
              operations: [operation],
            },
            { undefinedReferences: "refuse" },
          );
          expect(result.applied).toEqual([]);
          expect(result.undoHandle).toBeNull();
          expect(result.skipped).toMatchObject([
            {
              id: "replace",
              reason: "unsupportedBlock",
              canonicalRefusal: { gap: "comment-replacement-alignment" },
            },
          ]);
          expect(reviewer.toDocument()).toEqual(before);
          expect(comments(reviewer)).toEqual(original);
          expect(new Uint8Array(await reviewer.toBuffer())).toEqual(
            new Uint8Array(saved),
          );
        }
        return;
      }
      apply(direct, operation, "direct");
      const result = apply(tracked, operation, "tracked-changes");
      const pending = await tracked.toBuffer();
      expect(tracked.acceptAll().skipped).toEqual([]);
      expect(comments(direct)).toEqual(comments(legacy));
      expect(comments(tracked)).toEqual(comments(legacy));
      for (const reviewer of [direct, tracked])
        expect(
          comments(await open(await reviewer.toBuffer(), "canonical")),
        ).toEqual(comments(legacy));
      const rejected = await open(pending, "canonical");
      expect(rejected.rejectAll().skipped).toEqual([]);
      expect(comments(rejected)).toEqual(original);
      for (const decision of ["accept", "reject"] as const) {
        for (const order of ["forward", "reverse"] as const) {
          const individually = await open(pending, "canonical");
          const changes = individually.getChanges();
          for (const change of order === "forward"
            ? changes
            : changes.toReversed())
            expect(
              decision === "accept"
                ? individually.acceptChange(change)
                : individually.rejectChange(change),
            ).toBe(true);
          const expected = decision === "accept" ? comments(legacy) : original;
          expect(comments(individually)).toEqual(expected);
          expect(
            comments(await open(await individually.toBuffer(), "canonical")),
          ).toEqual(expected);
        }
      }
      const undone = await open(bytes, "canonical");
      const undoResult = apply(undone, operation, "tracked-changes");
      if (undoResult.undoHandle === null || result.undoHandle === null)
        throw new Error("Replacement must have an undo handle");
      expect(undone.undoDocumentOperations(undoResult.undoHandle).status).toBe(
        "undone",
      );
      expect(comments(undone)).toEqual(original);
    });
  }
});
