/**
 * File-system primitives shared by the document reader and the sidecar
 * files. Kept free of the document model so `editor-lease.ts` (which an
 * editor host bundles) does not pull in the parser.
 */

import { createHash } from "node:crypto";
import { constants } from "node:fs";

/** A file's version: the lowercase hex SHA-256 of its bytes. */
export const fileVersionOf = (bytes: Uint8Array): string =>
  createHash("sha256").update(bytes).digest("hex");

/** Which file a path named when it was read: device and inode. */
export type FileIdentity = { dev: number; ino: number };

export const sameFile = (left: FileIdentity, right: FileIdentity): boolean =>
  left.dev === right.dev && left.ino === right.ino;

/**
 * `O_NOFOLLOW` where the platform has it: opening a path whose last
 * component is a symlink fails instead of following it.
 */
export const NO_FOLLOW: number = constants.O_NOFOLLOW ?? 0;

/** The `code` of a Node.js file-system error, when there is one. */
export const errnoCode = (error: unknown): string | undefined => {
  if (typeof error !== "object" || error === null || !("code" in error)) {
    return undefined;
  }
  return typeof error.code === "string" ? error.code : undefined;
};
