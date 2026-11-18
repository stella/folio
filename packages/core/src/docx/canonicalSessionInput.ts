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
import { ensureUniqueRevisionIds } from "./ensureUniqueRevisionIds";
import { parseDocx } from "./parser";

const CanonicalDocxInputErrorBase: TaggedErrorClass<"CanonicalDocxInputError"> =
  TaggedError("CanonicalDocxInputError");

export class CanonicalDocxInputError extends CanonicalDocxInputErrorBase<{
  message: string;
  cause?: unknown;
}> {}

/**
 * Canonical seeding supports plaintext packages; refuse encryption before ZIP
 * normalization. A package that repeats a paragraph, content control or
 * revision id is given unique ones, every paragraph, control and revision
 * kept.
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
    // Paragraph ids, content control ids and revision ids, each unique in
    // the package as the model holds them.
    try: async () =>
      (
        await ensureUniqueRevisionIds(
          (
            await ensureUniqueControlIds(
              (await ensureParaIds(bytes.value)).docx,
            )
          ).docx,
        )
      ).docx,
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
 * Whether no two revision records of the model carry one id: a tracked
 * change or property change names itself by its `info.id`, as the package's
 * `w:id`.
 */
const holdsUniqueRevisionIds = (document: Document): boolean => {
  const seen = new Set<number>();
  // Every occurrence counts, as the package census counts it: a record held
  // in two places is two records. Only the section view, which restates the
  // body's paragraphs, is left out, as the census leaves it out.
  const open = new Set<object>();
  const visit = (value: unknown): boolean => {
    if (typeof value !== "object" || value === null || open.has(value))
      return true;
    open.add(value);
    const unique = visitRecord(value);
    open.delete(value);
    return unique;
  };
  const visitRecord = (value: object): boolean => {
    if (value instanceof Map) return [...value.values()].every(visit);
    if (Array.isArray(value)) return value.every(visit);
    if (
      "info" in value &&
      typeof value.info === "object" &&
      value.info !== null &&
      "id" in value.info &&
      typeof value.info.id === "number" &&
      "author" in value.info &&
      typeof value.info.author === "string"
    ) {
      if (seen.has(value.info.id)) return false;
      seen.add(value.info.id);
    }
    return Object.values(value).every(visit);
  };
  return visit({
    ...document.package,
    document: { ...document.package.document, sections: undefined },
  });
};

/**
 * Whether a model already holds the ids a canonical session addresses its
 * paragraphs and revisions by, each once.
 */
export const holdsUniqueSeedIds = (document: Document): boolean =>
  holdsUniqueParagraphIds(document) && holdsUniqueRevisionIds(document);

/**
 * A host's document model, ready to seed a canonical session. A model whose
 * paragraphs and revisions already hold unique ids opens as it is. Any other is written to
 * its package and read back through {@link prepareCanonicalDocxInput}, so its
 * paragraphs get exactly the ids that package gets when it is opened as
 * bytes: one owner establishes ids for both entry points.
 */
export const prepareCanonicalDocumentInput = async (
  document: Document,
): Promise<
  Err<Document, CanonicalDocxInputError> | Ok<Document, CanonicalDocxInputError>
> => {
  if (holdsUniqueSeedIds(document)) return Result.ok(document);
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
