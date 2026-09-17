import { DRAWING_RAW_XML_MODES } from "@stll/docx-core/model";

import type { DrawingContent, DrawingRawXmlMode, Image } from "../types/document";
import { canonicalJson } from "../utils/canonicalJson";

/**
 * Whether the editor may manipulate the modeled image of a classified drawing.
 *
 * A mode exists precisely because `rawXml` says more than the model does, so
 * both current modes refuse: a resize would either serialize a placeholder
 * (preserve-only) or one child picture in place of a group (preview-only).
 * Totality is the point — a third mode cannot be added without deciding here.
 */
const DRAWING_RAW_XML_MODE_ALLOWS_DIRECT_EDIT = {
  [DRAWING_RAW_XML_MODES.PRESERVE_ONLY]: false,
  [DRAWING_RAW_XML_MODES.PREVIEW_ONLY]: false,
} as const satisfies Record<DrawingRawXmlMode, boolean>;

/** An unclassified drawing is an ordinary editable projection; a classified one is not. */
export const allowsDirectDrawingEdit = (mode: DrawingRawXmlMode | undefined): boolean =>
  mode === undefined || DRAWING_RAW_XML_MODE_ALLOWS_DIRECT_EDIT[mode];

/** Narrow an unvalidated value (a ProseMirror attr) to a raw-XML mode. */
export const isDrawingRawXmlMode = (value: unknown): value is DrawingRawXmlMode =>
  typeof value === "string" && value in DRAWING_RAW_XML_MODE_ALLOWS_DIRECT_EDIT;

/**
 * Stands in for a preview's fingerprint once the editor has changed the image
 * it renders. `canonicalJson` always produces an object literal, so this can
 * never collide with a real fingerprint and the drawing can never replay.
 */
export const EDITED_PREVIEW_FINGERPRINT = "editedPreview";

const editableImageProjection = ({
  id: _id,
  rId: _rId,
  src: _src,
  mimeType: _mimeType,
  filename: _filename,
  ...image
}: Image) => image;

/** Fingerprints modeled image fields which make raw DrawingML stale when edited. */
export const imageRawXmlFingerprint = (image: Image): string =>
  canonicalJson(editableImageProjection(image));

/** Editable raw drawing XML can replay only while its modeled projection is unchanged. */
export const canReplayEditableImageRawXml = (drawing: DrawingContent): boolean => {
  if (drawing.rawXmlMode === DRAWING_RAW_XML_MODES.PRESERVE_ONLY) {
    return true;
  }
  return (
    drawing.rawImageFingerprint === undefined ||
    drawing.rawImageFingerprint === imageRawXmlFingerprint(drawing.image)
  );
};

/**
 * What a drawing costs the document when Folio writes it back.
 *
 * `native` and `replayable` both survive a save intact, so neither restricts
 * editing; only `opaque` loses content.
 */
export const DRAWING_SAFETY_CLASSES = {
  /** No raw XML: Folio owns the whole drawing and regenerates it from the model. */
  NATIVE: "native",
  /** Raw XML the serializer replays verbatim, so every unmodeled attribute survives. */
  REPLAYABLE: "replayable",
  /** Raw XML the serializer can neither replay nor regenerate without losing the media. */
  OPAQUE: "opaque",
} as const;

export type DrawingSafetyClass =
  (typeof DRAWING_SAFETY_CLASSES)[keyof typeof DRAWING_SAFETY_CLASSES];

/**
 * Regenerated DrawingML points at `image.rId`, and `serializePicGraphic` falls
 * back to `"rId1"` when that id is empty, which rebinds the picture to whichever
 * relationship happens to be first. A drawing with no relationship id therefore
 * has no faithful regeneration.
 */
const canRegenerateDrawing = (drawing: DrawingContent): boolean => drawing.image.rId !== "";

/**
 * Classify a drawing by what the run serializer will do with it.
 *
 * Shares {@link canReplayEditableImageRawXml} with the serializer so a
 * compatibility probe and a save can never disagree about the same drawing.
 */
export const classifyDrawingSafety = (drawing: DrawingContent): DrawingSafetyClass => {
  if (drawing.rawXml === undefined) {
    return DRAWING_SAFETY_CLASSES.NATIVE;
  }
  if (canReplayEditableImageRawXml(drawing)) {
    return DRAWING_SAFETY_CLASSES.REPLAYABLE;
  }
  // A raster preview has no faithful regeneration: `image` is a render of the
  // group, and its `rId` is whichever child blip the rasterizer saw first, so
  // regenerating would emit that one picture in place of the whole group.
  if (drawing.rawXmlMode === DRAWING_RAW_XML_MODES.PREVIEW_ONLY) {
    return DRAWING_SAFETY_CLASSES.OPAQUE;
  }
  return canRegenerateDrawing(drawing)
    ? DRAWING_SAFETY_CLASSES.NATIVE
    : DRAWING_SAFETY_CLASSES.OPAQUE;
};
