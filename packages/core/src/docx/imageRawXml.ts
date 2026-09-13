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
