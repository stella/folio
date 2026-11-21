/** One parser-fuzz case, isolated so the parent can kill a stalled parse. */

import assert from "node:assert/strict";

import { FolioDocxReviewer } from "../../ai-edits/headless";
import type { HeadlessReviewerSession } from "../../types/canonicalCapabilities";
import { CanonicalDocxInputError } from "../canonicalSessionInput";
import { DocxEncryptionError } from "../encryption/errors";
import { DocxParseError, parseDocx } from "../parser";

const publicRead = (reviewer: FolioDocxReviewer) => ({
  blocks: reviewer.getContent().map((block) => {
    if (block.kind === "diagnostic") {
      return {
        kind: block.kind,
        text: block.text,
        diagnostic: block.diagnostic,
      };
    }
    return {
      kind: block.kind,
      text: block.text,
      headingLevel: block.headingLevel,
      displayLabel: block.displayLabel,
      statedNumbering: block.statedNumbering,
    };
  }),
  stories: reviewer.listStories().map(({ handle, text }) => ({ handle, text })),
  notes: reviewer.getNotesAsText(),
  comments: reviewer
    .getComments()
    .map(({ author, text, anchoredText, done, replies }) => ({
      author,
      text,
      anchoredText,
      done,
      replies: replies.map((reply) => reply.text),
    })),
  changes: reviewer.getChanges().map(({ type, author }) => ({ type, author })),
});

const isProgrammerError = (error: unknown): boolean => {
  if (!(error instanceof Error)) return false;
  if (
    error instanceof TypeError ||
    error instanceof RangeError ||
    error instanceof ReferenceError
  ) {
    return true;
  }
  return "cause" in error && isProgrammerError(error.cause);
};

const isTypedRefusal = (error: unknown): boolean =>
  (error instanceof DocxParseError || error instanceof DocxEncryptionError) &&
  !isProgrammerError(error);

/** A canonical session refuses what it cannot hold with a typed error. */
const isTypedSessionRefusal = (error: unknown): boolean =>
  (error instanceof CanonicalDocxInputError ||
    (error instanceof Error &&
      "_tag" in error &&
      error._tag === "CanonicalSessionError")) &&
  !isProgrammerError(error);

const run = async (): Promise<void> => {
  const encoded = process.argv.at(2);
  if (!encoded) throw new Error("Missing parser mutation case bytes");
  const SESSIONS = [
    "legacy",
    "canonical",
  ] as const satisfies readonly HeadlessReviewerSession[];
  const session = SESSIONS.find(
    (candidate) => candidate === process.argv.at(3),
  );
  if (session === undefined) throw new Error("Missing reviewer session");
  const input = Uint8Array.from(Buffer.from(encoded, "base64")).buffer;
  try {
    await parseDocx(input, { preloadFonts: false, detectVariables: false });
  } catch (error) {
    assert.ok(
      isTypedRefusal(error),
      `untyped parser refusal: ${String(error)}`,
    );
    process.stdout.write("refused\n");
    return;
  }

  let reviewer: FolioDocxReviewer;
  try {
    reviewer = await FolioDocxReviewer.fromBuffer(input, { session });
  } catch (error) {
    assert.ok(
      isTypedSessionRefusal(error),
      `untyped session refusal: ${String(error)}`,
    );
    process.stdout.write("refused\n");
    return;
  }
  const before = publicRead(reviewer);
  const saved = await reviewer.toBuffer();
  const reopened = await FolioDocxReviewer.fromBuffer(saved, { session });
  assert.deepEqual(
    publicRead(reopened),
    before,
    "save/reopen changed public reader output",
  );
  process.stdout.write("opened\n");
};

await run().catch((error: unknown) => {
  process.stderr.write(
    error instanceof Error ? (error.stack ?? error.message) : String(error),
  );
  process.exitCode = 1;
});
