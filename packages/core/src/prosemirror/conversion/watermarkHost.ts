import type { BlockContent } from "../../types/document";
import { headerFooterToProseDoc, type ToProseDocOptions } from "./toProseDoc";
import { markDetachedWatermarkHost } from "./watermarkHostMarker";

type WatermarkHostConversionOptions = {
  content: BlockContent[];
  hostBlockIndex: number;
  options?: ToProseDocOptions;
};

/** Internal conversion path that marks the retained host of a detached watermark. */
export const headerFooterToProseDocWithDetachedWatermarkHost = ({
  content,
  hostBlockIndex,
  options,
}: WatermarkHostConversionOptions) => {
  const markedContent = content.map((block, blockIndex) =>
    blockIndex === hostBlockIndex ? markDetachedWatermarkHost(block) : block,
  );
  return headerFooterToProseDoc(markedContent, options);
};
