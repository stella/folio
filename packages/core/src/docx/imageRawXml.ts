import { DRAWING_RAW_XML_MODES } from "@stll/docx-core/model";

import type { DrawingContent, Image } from "../types/document";
import { canonicalJson } from "../utils/canonicalJson";

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
  return canRegenerateDrawing(drawing)
    ? DRAWING_SAFETY_CLASSES.NATIVE
    : DRAWING_SAFETY_CLASSES.OPAQUE;
};
