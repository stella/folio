import {
  type Err,
  type Ok,
  Result,
  TaggedError,
  type TaggedErrorClass,
} from "better-result";

import { type DocxInput, toArrayBuffer } from "../utils/docxInput";
import {
  detectDocxContainerType,
  DOCX_CONTAINER_TYPES,
} from "./encryption/containerFormat";
import { ensureParaIds } from "./ensureParaIds";
import { ensureUniqueControlIds } from "./ensureUniqueControlIds";

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
