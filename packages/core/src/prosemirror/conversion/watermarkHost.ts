import { panic } from "better-result";
import type { Node as PMNode } from "prosemirror-model";

import type { BlockContent } from "../../types/document";
import { headerFooterToProseDocInternal, type ToProseDocOptions } from "./toProseDoc";

const DETACHED_WATERMARK_HOST_ATTR = "_detachedWatermarkHost";

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

/** Read the internal host marker without exposing it through public paragraph attrs. */
export const expectDetachedWatermarkHostAttr = (node: PMNode): boolean => {
  if (node.type.name !== "paragraph") {
    panic(
      `Invalid ProseMirror detached watermark host attrs:\nnode.type.name: Expected paragraph, got ${node.type.name}.`,
    );
  }

  const value = Reflect.get(node.attrs, DETACHED_WATERMARK_HOST_ATTR);
  if (value === null || value === undefined) {
    return false;
  }
  if (typeof value !== "boolean") {
    panic(
      "Invalid ProseMirror detached watermark host attrs:\nparagraph.attrs._detachedWatermarkHost: Expected a boolean.",
    );
  }
  return value;
};
