import type { DrawingContent, RunContent } from "../types/content";

const SYNTHETIC_IMAGE_RID_PREFIX = "rId_img_";

/**
 * A drawing needs a new package image part only when it is model-driven.
 *
 * Raw OOXML drawings can carry data-URL previews for browser rendering, but
 * save replays their raw XML instead of serializing the preview image.
 */
export const isNewDataUrlDrawing = (content: RunContent): content is DrawingContent =>
  content.type === "drawing" &&
  !content.rawXml &&
  content.image.src?.startsWith("data:") === true &&
  (!content.image.rId || content.image.rId.startsWith(SYNTHETIC_IMAGE_RID_PREFIX));
