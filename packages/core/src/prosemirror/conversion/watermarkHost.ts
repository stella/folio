import type { BlockContent } from "../../types/document";
import { headerFooterToProseDocInternal, type ToProseDocOptions } from "./toProseDoc";

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
}: WatermarkHostConversionOptions) =>
  headerFooterToProseDocInternal({
    content,
    options,
    detachedWatermarkHostBlockIndex: hostBlockIndex,
  });
