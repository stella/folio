import { panic } from "better-result";

import type { FolioContentBlock } from "../../compare/content-types";

const OWNED_CONTENT_SNAPSHOT_BRAND: unique symbol = Symbol("owned-content-snapshot");

/**
 * An internal, already-canonical content projection.
 *
 * Public comparison inputs never enter this path. The payload is retained in
 * a closure-private WeakMap, so a structural lookalike cannot bypass the
 * adversarial input capture performed by `compareContent`.
 */
export type OwnedContentSnapshot = {
  readonly [OWNED_CONTENT_SNAPSHOT_BRAND]: true;
};

const blocksBySnapshot = new WeakMap<object, readonly FolioContentBlock[]>();

const freezeRecursively = (value: unknown): void => {
  if (typeof value !== "object" || value === null || Object.isFrozen(value)) return;
  for (const child of Object.values(value)) freezeRecursively(child);
  Object.freeze(value);
};

/**
 * @internal Take ownership of a freshly constructed trusted projection.
 *
 * The caller must not reuse the mutable builder after this call. Deep-freezing
 * in place closes every alias while retaining the exact canonical objects that
 * nominal DOCX operands bind to; cloning here would add a second full-story
 * allocation before the shared comparison budget has measured the projection.
 */
export const ownContentSnapshot = (
  blocks: readonly FolioContentBlock[],
): OwnedContentSnapshot => {
  freezeRecursively(blocks);
  const snapshot = Object.freeze({
    [OWNED_CONTENT_SNAPSHOT_BRAND]: true as const,
  });
  blocksBySnapshot.set(snapshot, blocks);
  return snapshot;
};

/** @internal Read the payload only from a capsule created in this module. */
export const ownedContentSnapshotBlocks = (
  snapshot: unknown,
): readonly FolioContentBlock[] | null => {
  if (typeof snapshot !== "object" || snapshot === null) return null;
  return blocksBySnapshot.get(snapshot) ?? null;
};

/** @internal Require the payload after a nominal type has crossed a module seam. */
export const requireOwnedContentSnapshotBlocks = (
  snapshot: OwnedContentSnapshot,
): readonly FolioContentBlock[] =>
  blocksBySnapshot.get(snapshot) ?? panic("An owned content snapshot was not created by Folio");
