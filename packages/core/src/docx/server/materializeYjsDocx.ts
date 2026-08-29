import { Result, TaggedError } from "better-result";
import { initProseMirrorDoc } from "y-prosemirror";
import * as Y from "yjs";

import { fromProseDoc } from "../../prosemirror/conversion/fromProseDoc";
import { schema } from "../../prosemirror/schema";
import { parseDocx } from "../parser";
import { repackDocx } from "../rezip";

/** Yjs fragment that stores Folio's canonical ProseMirror document. */
export const FOLIO_YJS_PROSEMIRROR_FRAGMENT_NAME = "prosemirror";

/** Maximum accepted size of a complete Yjs collaboration state update. */
export const FOLIO_YJS_UPDATE_MAX_BYTES = 10 * 1024 * 1024;

/** Stable failure codes returned by server-side Yjs-to-DOCX materialization. */
export const FOLIO_YJS_DOCX_MATERIALIZATION_ERROR_CODES = [
  "empty_update",
  "invalid_update",
  "missing_document",
  "update_too_large",
] as const;

/** Failure code for a rejected Yjs-to-DOCX materialization request. */
export type FolioYjsDocxMaterializationErrorCode =
  (typeof FOLIO_YJS_DOCX_MATERIALIZATION_ERROR_CODES)[number];

/** Typed failure raised when a collaboration snapshot cannot be materialized. */
export class FolioYjsDocxMaterializationError extends TaggedError(
  "FolioYjsDocxMaterializationError",
)<{
  code: FolioYjsDocxMaterializationErrorCode;
  message: string;
  cause?: unknown;
}> {}

/** Inputs for materializing a complete Yjs state update into a DOCX package. */
export type MaterializeYjsDocxOptions = {
  /** Original DOCX whose package parts and non-body stories must be preserved. */
  sourceDocx: ArrayBuffer | Uint8Array;
  /** Complete Yjs state update containing Folio's ProseMirror fragment. */
  yjsUpdate: Uint8Array;
};

const readProseMirrorDocument = (yjsUpdate: Uint8Array) => {
  if (yjsUpdate.byteLength === 0) {
    throw new FolioYjsDocxMaterializationError({
      code: "empty_update",
      message: "Cannot materialize DOCX from an empty Yjs update.",
    });
  }
  if (yjsUpdate.byteLength > FOLIO_YJS_UPDATE_MAX_BYTES) {
    throw new FolioYjsDocxMaterializationError({
      code: "update_too_large",
      message: "Yjs update exceeds the DOCX materialization limit.",
    });
  }

  const ydoc = new Y.Doc();
  const parsed = Result.try({
    try: () => {
      Y.applyUpdate(ydoc, yjsUpdate);
      const fragment = ydoc.getXmlFragment(FOLIO_YJS_PROSEMIRROR_FRAGMENT_NAME);
      if (fragment.length === 0) {
        throw new FolioYjsDocxMaterializationError({
          code: "missing_document",
          message: "Yjs update does not contain a Folio document.",
        });
      }
      return initProseMirrorDoc(fragment, schema).doc;
    },
    catch: (cause) =>
      cause instanceof FolioYjsDocxMaterializationError
        ? cause
        : new FolioYjsDocxMaterializationError({
            code: "invalid_update",
            message: "Yjs update is not a valid Folio collaboration snapshot.",
            cause,
          }),
  });
  ydoc.destroy();
  if (parsed.isOk()) {
    return parsed.value;
  }
  throw parsed.error;
};

/**
 * Materialize a complete Folio Yjs state update into a DOCX while preserving
 * package parts from the source document. This is the server-side equivalent
 * of the browser editor's full save path for the main document story.
 */
export const materializeYjsDocx = async ({
  sourceDocx,
  yjsUpdate,
}: MaterializeYjsDocxOptions): Promise<ArrayBuffer> => {
  const proseMirrorDocument = readProseMirrorDocument(yjsUpdate);
  const baseDocument = await parseDocx(sourceDocx, { preloadFonts: false });
  return await repackDocx(fromProseDoc(proseMirrorDocument, baseDocument));
};
