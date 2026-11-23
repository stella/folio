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
  paragraphsThroughTextBoxes,
} from "@stll/docx-core/ops";
import { MalformedXmlReferenceError } from "@stll/docx-core/xml";

import type { Document } from "../types/document";
import { type DocxInput, toArrayBuffer } from "../utils/docxInput";
import {
  detectDocxContainerType,
  DOCX_CONTAINER_TYPES,
} from "./encryption/containerFormat";
import { ensureParaIds } from "./ensureParaIds";
import { ensureUniqueControlIds } from "./ensureUniqueControlIds";
import { ensureUniqueRevisionIds } from "./ensureUniqueRevisionIds";
import { keepNormalizedSourceParts } from "./normalizedSourceParts";
import { MalformedPackageRelationshipsError } from "./packageRelationshipsRoot";
import { parseDocx, type ParseOptions } from "./parser";

const CanonicalDocxInputErrorBase: TaggedErrorClass<"CanonicalDocxInputError"> =
  TaggedError("CanonicalDocxInputError");

export class CanonicalDocxInputError extends CanonicalDocxInputErrorBase<{
  message: string;
  cause?: unknown;
}> {}

/** Why a package cannot seed a canonical session. */
export type CanonicalDocxInputFailure =
  | CanonicalDocxInputError
  | MalformedXmlReferenceError
  | MalformedPackageRelationshipsError;

/**
 * The malformation a normalization step refused a part for, if that is why:
 * a reference XML refuses, or package relationships that are not. Both
 * engines refuse a package with these errors.
 */
const malformedPartCause = (
  cause: unknown,
):
  | MalformedXmlReferenceError
  | MalformedPackageRelationshipsError
  | undefined => {
  for (let current = cause; current instanceof Error; current = current.cause) {
    if (
      current instanceof MalformedXmlReferenceError ||
      current instanceof MalformedPackageRelationshipsError
    )
      return current;
  }
  return undefined;
};

/**
 * Canonical seeding supports plaintext packages; refuse encryption before ZIP
 * normalization. A package that repeats a paragraph, content control or
 * revision id is given unique ones, every paragraph, control and revision
 * kept.
 */
export const prepareCanonicalDocxInput = async (
  input: DocxInput,
): Promise<
  | Err<ArrayBuffer, CanonicalDocxInputFailure>
  | Ok<Uint8Array<ArrayBufferLike>, CanonicalDocxInputFailure>
> => {
  const source = await prepareCanonicalDocxSource(input);
  return source.isErr()
    ? Result.err(source.error)
    : Result.ok(source.value.docx);
};

/** A canonical input package and the paragraph ids normalizing it minted. */
export type CanonicalDocxSource = {
  docx: Uint8Array<ArrayBufferLike>;
  /**
   * Ids written for paragraphs the package names none for, or names one
   * another paragraph already holds: they say nothing about the paragraph
   * across versions of the document.
   */
  mintedParaIds: ReadonlySet<string>;
};

/** {@link prepareCanonicalDocxInput}, with the paragraph ids it minted. */
export const prepareCanonicalDocxSource = async (
  input: DocxInput,
): Promise<Result<CanonicalDocxSource, CanonicalDocxInputFailure>> => {
  const bytes = await Result.tryPromise({
    try: () => toArrayBuffer(input),
    catch: (cause) =>
      new CanonicalDocxInputError({
        message: "Cannot read canonical document bytes.",
        cause,
      }),
  });
  if (bytes.isErr()) return Result.err(bytes.error);
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
    try: async () => {
      const paragraphs = await ensureParaIds(bytes.value);
      const { docx } = await ensureUniqueRevisionIds(
        (await ensureUniqueControlIds(paragraphs.docx)).docx,
      );
      return { docx, mintedParaIds: new Set(paragraphs.mintedParaIds) };
    },
    catch: (cause): CanonicalDocxInputFailure =>
      malformedPartCause(cause) ??
      new CanonicalDocxInputError({
        message:
          cause instanceof Error
            ? cause.message
            : "Canonical document normalization failed.",
        cause,
      }),
  });
};

/** A package read for a canonical session. */
export type CanonicalDocx = {
  document: Document;
  /** The paragraph ids normalizing the package minted. */
  mintedParaIds: ReadonlySet<string>;
};

/**
 * How an editor parses a package it opens in a canonical session: no
 * template-variable scan, fonts loaded after the first render.
 */
export const CANONICAL_DOCX_PARSE_OPTIONS = {
  detectVariables: false,
  preloadFonts: false,
} as const;

/**
 * Read a package for a canonical session: normalize it
 * ({@link prepareCanonicalDocxSource}) and parse the result. A save of the
 * document writes each part the session leaves as read with the package's
 * own bytes, so normalizing changes no part an edit does not reach.
 */
export const openCanonicalDocx = async (
  input: DocxInput,
  options: ParseOptions = {},
): Promise<Result<CanonicalDocx, CanonicalDocxInputFailure>> => {
  const bytes = await Result.tryPromise({
    try: () => toArrayBuffer(input),
    catch: (cause) =>
      new CanonicalDocxInputError({
        message: "Cannot read canonical document bytes.",
        cause,
      }),
  });
  if (bytes.isErr()) return Result.err(bytes.error);
  const prepared = await prepareCanonicalDocxSource(bytes.value);
  if (prepared.isErr()) return Result.err(prepared.error);
  const document = await parseDocx(prepared.value.docx, options);
  await keepNormalizedSourceParts({ source: bytes.value, document });
  return Result.ok({ document, mintedParaIds: prepared.value.mintedParaIds });
};

/**
 * Whether each story paragraph holds a paraId and no two paragraphs hold the
 * same one: the ids a canonical session addresses paragraphs by.
 */
export const holdsUniqueParagraphIds = (document: Document): boolean => {
  // A text box's paragraphs are addressed through the story drawing it.
  const paragraphs = documentStories(document).flatMap((story) =>
    paragraphsThroughTextBoxes(document, story),
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
  | Err<Document, CanonicalDocxInputFailure>
  | Ok<Document, CanonicalDocxInputFailure>
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
