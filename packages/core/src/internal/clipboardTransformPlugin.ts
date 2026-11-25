import { Plugin } from "prosemirror-state";

/** Retain the copy and paste transforms a view can use independently of an editor plugin. */
export const clipboardTransformPlugin = ({
  transformPastedHTML,
  transformPasted,
  transformCopied,
  clipboardTextSerializer,
}: Plugin["props"]) => {
  if (
    transformPastedHTML === undefined &&
    transformPasted === undefined &&
    transformCopied === undefined &&
    clipboardTextSerializer === undefined
  )
    return [];

  return [
    new Plugin({
      props: {
        ...(transformPastedHTML === undefined ? {} : { transformPastedHTML }),
        ...(transformPasted === undefined ? {} : { transformPasted }),
        ...(transformCopied === undefined ? {} : { transformCopied }),
        ...(clipboardTextSerializer === undefined
          ? {}
          : { clipboardTextSerializer }),
      },
    }),
  ];
};
