import { TaggedError, type TaggedErrorClass } from "better-result";

/**
 * Why an archive or entry was refused.
 *
 * - `malformed`: the records cannot be read, or disagree with each other.
 * - `unsupported`: encryption, a compression method other than STORE or
 *   DEFLATE, or a multi-disk archive.
 * - `invalid-entry-name`: an absolute, parent-relative, backslash, NUL or
 *   non-UTF-8 entry name.
 * - `duplicate-entry`: two records name the same entry.
 * - `too-many-entries`, `entry-too-large`, `total-too-large`,
 *   `compression-ratio`: a read limit.
 * - `declared-size-exceeded`: entry data inflates past its declared size;
 *   inflation stops there.
 * - `corrupt-entry`: entry data does not inflate, inflates short of its
 *   declared size, or fails its CRC-32.
 */
export type ZipErrorReason =
  | "malformed"
  | "unsupported"
  | "invalid-entry-name"
  | "duplicate-entry"
  | "too-many-entries"
  | "entry-too-large"
  | "total-too-large"
  | "compression-ratio"
  | "declared-size-exceeded"
  | "corrupt-entry";

const ZipErrorBase: TaggedErrorClass<"ZipError"> = TaggedError("ZipError");

/** A ZIP archive or entry the module refuses to read or write. */
export class ZipError extends ZipErrorBase<{
  message: string;
  reason: ZipErrorReason;
  /** The entry the error concerns, when it concerns one. */
  path?: string;
  cause?: unknown;
}> {}
