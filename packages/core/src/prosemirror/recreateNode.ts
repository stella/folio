import type { Fragment, Mark, Node as PMNode } from "prosemirror-model";

type RecreateProseNodeOptions = {
  attrs?: PMNode["attrs"];
  content?: Fragment | PMNode | readonly PMNode[] | null;
  marks?: readonly Mark[];
};

/** Immutable ProseMirror rebuild with no hidden ownership side channel. */
export const recreateProseNode = (
  source: PMNode,
  options: RecreateProseNodeOptions = {},
): PMNode =>
  source.type.create(
    options.attrs ?? source.attrs,
    options.content === undefined ? source.content : options.content,
    options.marks ?? source.marks,
  );
