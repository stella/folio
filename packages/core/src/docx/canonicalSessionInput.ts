import { Result, TaggedError } from "better-result";

import { ensureParaIds } from "./ensureParaIds";
import { DOCX_CONTAINER_TYPES, detectDocxContainerType } from "./encryption/containerFormat";
import { toArrayBuffer, type DocxInput } from "../utils/docxInput";

export class CanonicalDocxInputError extends TaggedError("CanonicalDocxInputError")<{
  message: string;
  cause?: unknown;
}> {}

/** Canonical seeding supports plaintext packages; refuse encryption before ZIP normalization. */
export const prepareCanonicalDocxInput = async (input: DocxInput) => {
  const bytes = await Result.tryPromise({
    try: () => toArrayBuffer(input),
    catch: (cause) =>
      new CanonicalDocxInputError({ message: "Cannot read canonical document bytes.", cause }),
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
    try: async () => (await ensureParaIds(bytes.value)).docx,
    catch: (cause) =>
      new CanonicalDocxInputError({
        message:
          cause instanceof Error ? cause.message : "Canonical document normalization failed.",
        cause,
      }),
  });
};
