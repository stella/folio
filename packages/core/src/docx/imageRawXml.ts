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

/** Whether the modeled image still says what the captured raw drawing says. */
const modelAgreesWithRawXml = (drawing: { image: Image; rawImageFingerprint?: string }): boolean =>
  drawing.rawImageFingerprint === undefined ||
  drawing.rawImageFingerprint === imageRawXmlFingerprint(drawing.image);

/**
 * Whether the serializer writes the captured raw XML back.
 *
 * For an unclassified drawing the capture is a cache of the model, so it
 * replays only while the two agree. For a classified one the capture *is* the
 * drawing: a preserve-only drawing has a placeholder image, and a preview-only
 * group has a raster render of shapes the model cannot hold. Regenerating
 * either writes something the source never contained — for the group, at best
 * the one child picture the rasterizer saw first under that child's
 * relationship, in place of the whole group. So a classified capture is always
 * written back; a stale fingerprint means the edit that made it stale is lost,
 * which {@link classifyDrawingSafety} reports as `opaque`, and never that
 * folio may build a replacement.
 */
export const canReplayEditableImageRawXml = (drawing: DrawingContent): boolean => {
  switch (drawing.rawXmlMode) {
    case DRAWING_RAW_XML_MODES.PRESERVE_ONLY:
    case DRAWING_RAW_XML_MODES.PREVIEW_ONLY:
      return true;
    case undefined:
      return modelAgreesWithRawXml(drawing);
    default: {
      const exhaustive: never = drawing;
      return exhaustive;
    }
  }
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
 * A drawing with no picture relationship has no picture to regenerate. The
 * serializer writes the anchor back without a graphic, which is faithful for an
 * anchor that never had one and lossy for a chart or an OLE frame, so the
 * drawing counts as opaque and an edit must block the save.
 */
const canRegenerateDrawing = (drawing: DrawingContent): boolean => drawing.image.rId !== undefined;

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
  switch (drawing.rawXmlMode) {
    case DRAWING_RAW_XML_MODES.PRESERVE_ONLY:
      return DRAWING_SAFETY_CLASSES.REPLAYABLE;
    // A raster preview has no faithful regeneration: `image` is a render of
    // the group, so the save replays the group and any edit to the render is
    // what is lost.
    case DRAWING_RAW_XML_MODES.PREVIEW_ONLY:
      return modelAgreesWithRawXml(drawing)
        ? DRAWING_SAFETY_CLASSES.REPLAYABLE
        : DRAWING_SAFETY_CLASSES.OPAQUE;
    case undefined:
      if (modelAgreesWithRawXml(drawing)) {
        return DRAWING_SAFETY_CLASSES.REPLAYABLE;
      }
      return canRegenerateDrawing(drawing)
        ? DRAWING_SAFETY_CLASSES.NATIVE
        : DRAWING_SAFETY_CLASSES.OPAQUE;
    default: {
      const exhaustive: never = drawing;
      return exhaustive;
    }
  }
};
