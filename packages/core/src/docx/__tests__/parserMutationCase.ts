/** One parser-fuzz case, isolated so the parent can kill a stalled parse. */

import assert from "node:assert/strict";

import { FolioDocxReviewer } from "../../ai-edits/headless";
import { DocxEncryptionError } from "../encryption/errors";
import { DocxParseError, parseDocx } from "../parser";

const publicRead = (reviewer: FolioDocxReviewer) => ({
  blocks: reviewer.getContent().map(({ kind, text, headingLevel, displayLabel, listLevel }) => ({
    kind,
    text,
    headingLevel,
    displayLabel,
    listLevel,
  })),
  stories: reviewer.listStories().map(({ handle, text }) => ({ handle, text })),
  notes: reviewer.getNotesAsText(),
  comments: reviewer.getComments().map(({ author, text, anchoredText, done, replies }) => ({
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

const run = async (): Promise<void> => {
  const encoded = process.argv.at(2);
  if (!encoded) throw new Error("Missing parser mutation case bytes");
  const input = Uint8Array.from(Buffer.from(encoded, "base64")).buffer;
  try {
    await parseDocx(input, { preloadFonts: false, detectVariables: false });
  } catch (error) {
    assert.ok(isTypedRefusal(error), `untyped parser refusal: ${String(error)}`);
    process.stdout.write("refused\n");
    return;
  }

  const reviewer = await FolioDocxReviewer.fromBuffer(input);
  const before = publicRead(reviewer);
  const saved = await reviewer.toBuffer();
  const reopened = await FolioDocxReviewer.fromBuffer(saved);
  assert.deepEqual(publicRead(reopened), before, "save/reopen changed public reader output");
  process.stdout.write("opened\n");
};

await run().catch((error: unknown) => {
  process.stderr.write(error instanceof Error ? (error.stack ?? error.message) : String(error));
  process.exitCode = 1;
});
