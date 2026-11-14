import {
  type Err,
  type Ok,
  Result,
  TaggedError,
  type TaggedErrorClass,
} from "better-result";

import {
  documentStories,
  idKey,
  packageParagraphIds,
  storyBody,
  storyParagraphs,
} from "@stll/docx-core/ops";

import type { Document } from "../types/document";
import { type DocxInput, toArrayBuffer } from "../utils/docxInput";
import {
  detectDocxContainerType,
  DOCX_CONTAINER_TYPES,
} from "./encryption/containerFormat";
import { ensureParaIds } from "./ensureParaIds";
import { ensureUniqueControlIds } from "./ensureUniqueControlIds";
import { parseDocx } from "./parser";

const CanonicalDocxInputErrorBase: TaggedErrorClass<"CanonicalDocxInputError"> =
  TaggedError("CanonicalDocxInputError");

export class CanonicalDocxInputError extends CanonicalDocxInputErrorBase<{
  message: string;
  cause?: unknown;
}> {}

/**
 * Canonical seeding supports plaintext packages; refuse encryption before ZIP
 * normalization. A package that repeats a paragraph or content control id is
 * given unique ones, every paragraph and control kept.
 */
export const prepareCanonicalDocxInput = async (
  input: DocxInput,
): Promise<
  | Err<ArrayBuffer, CanonicalDocxInputError>
  | Ok<Uint8Array<ArrayBufferLike>, CanonicalDocxInputError>
> => {
  const bytes = await Result.tryPromise({
    try: () => toArrayBuffer(input),
    catch: (cause) =>
      new CanonicalDocxInputError({
        message: "Cannot read canonical document bytes.",
        cause,
      }),
  });
  if (bytes.isErr()) return bytes;
  if (detectDocxContainerType(bytes.value) === DOCX_CONTAINER_TYPES.CFB) {
    return Result.err(
      new CanonicalDocxInputError({
        message:
          "Password-protected documents are unavailable in the experimental canonical session.",
      }),
    );
  }
  return Result.tryPromise({
    // Paragraph ids and content control ids, each unique in the package as
    // the model holds them.
    try: async () =>
      (await ensureUniqueControlIds((await ensureParaIds(bytes.value)).docx))
        .docx,
    catch: (cause) =>
      new CanonicalDocxInputError({
        message:
          cause instanceof Error
            ? cause.message
            : "Canonical document normalization failed.",
        cause,
      }),
  });
};

/**
 * Whether each story paragraph holds a paraId and no two paragraphs hold the
 * same one: the ids a canonical session addresses paragraphs by.
 */
export const holdsUniqueParagraphIds = (document: Document): boolean => {
  const paragraphs = documentStories(document).flatMap((story) =>
    storyParagraphs(storyBody(document, story)),
  );
  if (paragraphs.some(({ paragraph }) => paragraph.paraId === undefined))
    return false;
  const ids = packageParagraphIds(document.package).map(idKey);
  return new Set(ids).size === ids.length;
};

/**
 * A host's document model, ready to seed a canonical session. A model whose
 * paragraphs already hold unique ids opens as it is. Any other is written to
 * its package and read back through {@link prepareCanonicalDocxInput}, so its
 * paragraphs get exactly the ids that package gets when it is opened as
 * bytes: one owner establishes ids for both entry points.
 */
export const prepareCanonicalDocumentInput = async (
  document: Document,
): Promise<
  Err<Document, CanonicalDocxInputError> | Ok<Document, CanonicalDocxInputError>
> => {
  if (holdsUniqueParagraphIds(document)) return Result.ok(document);
  const written = await Result.tryPromise({
    try: async () => {
      const { createDocx } = await import("./rezip");
      return createDocx(document);
    },
    catch: (cause) =>
      new CanonicalDocxInputError({
        message:
          cause instanceof Error
            ? cause.message
            : "Cannot write the document to establish paragraph ids.",
        cause,
      }),
  });
  if (written.isErr()) return written;
  const prepared = await prepareCanonicalDocxInput(written.value);
  if (prepared.isErr()) return prepared;
  return Result.tryPromise({
    try: () =>
      parseDocx(prepared.value, {
        detectVariables: false,
        preloadFonts: false,
      }),
    catch: (cause) =>
      new CanonicalDocxInputError({
        message:
          cause instanceof Error
            ? cause.message
            : "Cannot read the document back with its paragraph ids.",
        cause,
      }),
  });
};
