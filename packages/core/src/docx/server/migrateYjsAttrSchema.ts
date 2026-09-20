import { Result, TaggedError } from "better-result";
import * as Y from "yjs";

import {
  FOLIO_YJS_ATTR_SCHEMA_VERSION,
  applyAttrSchemaMigrations,
  type FolioYjsAttrSchemaVersion,
  readYjsAttrSchemaVersion,
} from "../../prosemirror/yjsDocumentMetadata";
import {
  FOLIO_YJS_PROSEMIRROR_FRAGMENT_NAME,
  FOLIO_YJS_UPDATE_MAX_BYTES,
} from "./materializeYjsDocx";

/** Stable failure codes returned by an offline snapshot migration. */
export const FOLIO_YJS_SNAPSHOT_MIGRATION_ERROR_CODES = [
  "invalid_update",
  "unsupported_version",
  "update_too_large",
] as const;

/** Failure code for a snapshot that could not be migrated. */
export type FolioYjsSnapshotMigrationErrorCode =
  (typeof FOLIO_YJS_SNAPSHOT_MIGRATION_ERROR_CODES)[number];

/** Typed failure raised when a collaboration snapshot cannot be migrated. */
export class FolioYjsSnapshotMigrationError extends TaggedError("FolioYjsSnapshotMigrationError")<{
  code: FolioYjsSnapshotMigrationErrorCode;
  message: string;
  cause?: unknown;
}> {}

/** What an offline migration did to one snapshot. */
export type FolioYjsSnapshotMigrationResult = {
  /** The complete state update to store in place of the input. */
  update: Uint8Array;
  /** The attr-schema version the input was written under. */
  fromVersion: FolioYjsAttrSchemaVersion;
  /** The attr-schema version the output carries. */
  toVersion: FolioYjsAttrSchemaVersion;
  /** Paragraphs whose attrs the migration rewrote. */
  paragraphsRewritten: number;
};

/**
 * Carry a stored collaboration snapshot forward to the attr schema this build
 * writes, without a ProseMirror schema and without rebuilding a single node.
 *
 * A Yjs XML element's attributes are ordinary CRDT map entries holding JSON, so
 * {@link applyAttrSchemaMigrations} rewrites values in the fragment directly.
 * Nothing here can trip y-prosemirror's node rebuild, which deletes an element
 * it fails to construct. The function is pure and offline: no network, no
 * DOCX, no editor.
 *
 * A snapshot already at the current version is returned byte for byte, so a
 * host may run this over its whole corpus repeatedly.
 *
 * **The contract a reseed must respect.** This function never reseeds and never
 * touches the paragraph-source contract: it preserves whichever source binding
 * the snapshot already carries. A reseed does the opposite, and Folio holds it
 * to one rule: *the DOCX handed to the seeding editor must be the same bytes
 * the host will later pass to `materializeYjsDocx` as that room's source*.
 * The contract is a digest of those bytes, so seeding from one document while
 * recording another as the room's source makes every later save fail with
 * `source_mismatch`. The reverse mistake is worse and Folio cannot see it: a
 * host that reseeds from the document a room was *created* from, rather than
 * from the latest content it published, produces a perfectly well-formed
 * snapshot whose collaborative history has been replaced by an older document.
 * Rebind the recorded source in the same statement that clears the seed, or do
 * not clear it.
 */
export const migrateFolioYjsSnapshot = (
  update: Uint8Array,
): Result<FolioYjsSnapshotMigrationResult, FolioYjsSnapshotMigrationError> => {
  if (update.byteLength === 0) {
    return Result.err(
      new FolioYjsSnapshotMigrationError({
        code: "invalid_update",
        message: "Cannot migrate an empty Yjs update.",
      }),
    );
  }
  if (update.byteLength > FOLIO_YJS_UPDATE_MAX_BYTES) {
    return Result.err(
      new FolioYjsSnapshotMigrationError({
        code: "update_too_large",
        message: "Yjs update exceeds the snapshot migration limit.",
      }),
    );
  }

  const ydoc = new Y.Doc();
  const applied = Result.try({
    try: () => {
      Y.applyUpdate(ydoc, update);
    },
    catch: (cause) =>
      new FolioYjsSnapshotMigrationError({
        cause,
        code: "invalid_update",
        message: "Yjs update is not a valid Folio collaboration snapshot.",
      }),
  });
  if (applied.isErr()) {
    ydoc.destroy();
    return Result.err(applied.error);
  }

  const read = readYjsAttrSchemaVersion(ydoc);
  if (read.isErr()) {
    ydoc.destroy();
    return Result.err(
      new FolioYjsSnapshotMigrationError({
        cause: read.error,
        code: "unsupported_version",
        message: read.error.message,
      }),
    );
  }

  const fromVersion = read.value;
  if (fromVersion === FOLIO_YJS_ATTR_SCHEMA_VERSION) {
    ydoc.destroy();
    return Result.ok({
      fromVersion,
      paragraphsRewritten: 0,
      toVersion: FOLIO_YJS_ATTR_SCHEMA_VERSION,
      update,
    });
  }

  const paragraphsRewritten = applyAttrSchemaMigrations(
    ydoc,
    ydoc.getXmlFragment(FOLIO_YJS_PROSEMIRROR_FRAGMENT_NAME),
    fromVersion,
  );
  const migrated = Y.encodeStateAsUpdate(ydoc);
  ydoc.destroy();

  return Result.ok({
    fromVersion,
    paragraphsRewritten,
    toVersion: FOLIO_YJS_ATTR_SCHEMA_VERSION,
    update: migrated,
  });
};
