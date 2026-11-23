import { Result } from "better-result";

import {
  CANONICAL_DOCX_PARSE_OPTIONS,
  type CanonicalDocx,
  type CanonicalDocxInputFailure,
  openCanonicalDocx,
} from "../docx/canonicalSessionInput";
import type { Document } from "../types/document";
import { readModel } from "../types/documentModel";
import {
  type CanonicalSession,
  type CanonicalSessionError,
  createCanonicalSession,
} from "./canonicalSession";

/** A room's base snapshot identity: the SHA-256 of the DOCX bytes, in hex. */
export const canonicalSnapshotId = async (
  docx: ArrayBuffer | Uint8Array,
): Promise<string> => {
  const digest = await crypto.subtle.digest("SHA-256", new Uint8Array(docx));
  return [...new Uint8Array(digest)]
    .map((byte) => byte.toString(16).padStart(2, "0"))
    .join("");
};

/** A room's revision 0. */
export type CanonicalRoomBase = {
  /** The document the canonical session reads from the room's DOCX. */
  document: Document;
  /** `canonicalSnapshotId` of that DOCX. */
  snapshot: string;
};

export type CanonicalDocxSession = CanonicalDocx & {
  session: CanonicalSession;
};

/** Open `docx` in a canonical session, as an editor opens it. */
export const openCanonicalDocxSession = async (
  docx: ArrayBuffer,
  password?: string,
): Promise<
  Result<
    CanonicalDocxSession,
    CanonicalDocxInputFailure | CanonicalSessionError
  >
> => {
  const opened = await openCanonicalDocx(docx, {
    ...CANONICAL_DOCX_PARSE_OPTIONS,
    password,
  });
  if (opened.isErr()) return Result.err(opened.error);
  const { document } = opened.value;
  const session = createCanonicalSession(document, document.package.styles);
  if (session.isErr()) return Result.err(session.error);
  return Result.ok({ ...opened.value, session: session.value });
};

/**
 * A room's revision 0 for its authority: the document an editor's canonical
 * session reads from `docx`, and that DOCX's snapshot identity.
 */
export const openCanonicalRoomBase = async (
  docx: ArrayBuffer,
): Promise<
  Result<CanonicalRoomBase, CanonicalDocxInputFailure | CanonicalSessionError>
> => {
  const opened = await openCanonicalDocxSession(docx);
  if (opened.isErr()) return Result.err(opened.error);
  return Result.ok({
    document: readModel(opened.value.session.document),
    snapshot: await canonicalSnapshotId(docx),
  });
};
