import type { DrawingContent, RunContent } from "../types/content";

const SYNTHETIC_IMAGE_RID_PREFIX = "rId_img_";

/**
 * A drawing needs a new package image part only when it is model-driven.
 *
 * Raw OOXML drawings can carry previews for browser rendering. Only an explicit
 * synthetic relationship opts a detached raw picture into media registration.
 */
export const isNewDataUrlDrawing = (content: RunContent): content is DrawingContent =>
  content.type === "drawing" &&
  (!content.rawXml || content.image.rId?.startsWith(SYNTHETIC_IMAGE_RID_PREFIX) === true) &&
  content.image.src?.startsWith("data:") === true &&
  (!content.image.rId || content.image.rId.startsWith(SYNTHETIC_IMAGE_RID_PREFIX));
