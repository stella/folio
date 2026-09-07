import { panic } from "better-result";

const DETACHED_WATERMARK_HOST = Symbol("detachedWatermarkHost");
const DETACHED_WATERMARK_HOST_ATTR = "_detachedWatermarkHost";

export const markDetachedWatermarkHost = <Block extends { type: string }>(block: Block): Block => {
  if (block.type !== "paragraph") {
    return block;
  }
  const markedBlock = { ...block, [DETACHED_WATERMARK_HOST]: true };
  return markedBlock;
};

export const isDetachedWatermarkHost = <Block extends object>(block: Block): boolean =>
  Reflect.get(block, DETACHED_WATERMARK_HOST) === true;

/** Read the internal host marker without exposing it through public paragraph attrs. */
export const expectDetachedWatermarkHostAttr = (
  attrs: Readonly<Record<string, unknown>>,
): boolean => {
  const value = Reflect.get(attrs, DETACHED_WATERMARK_HOST_ATTR);
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
