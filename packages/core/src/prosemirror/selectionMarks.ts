import type { Mark, Node as PMNode } from "prosemirror-model";

const UNDERLINE_MARK = "underline";
const HIDDEN_UNDERLINE_STYLE = "none";

type CollectMarksInRangeOptions = {
  doc: PMNode;
  from: number;
  to: number;
};

const shouldReplaceMark = (current: Mark, incoming: Mark): boolean =>
  current.type.name === UNDERLINE_MARK &&
  current.attrs["style"] === HIDDEN_UNDERLINE_STYLE &&
  incoming.attrs["style"] !== HIDDEN_UNDERLINE_STYLE;

/** Collect one representative per mark type, preferring visible underline state. */
export const collectMarksInRange = ({ doc, from, to }: CollectMarksInRangeOptions): Mark[] => {
  const seen = new Map<string, Mark>();

  doc.nodesBetween(from, to, (node) => {
    if (!node.isText) {
      return;
    }

    for (const mark of node.marks) {
      const name = mark.type.name;
      const current = seen.get(name);
      if (!current || shouldReplaceMark(current, mark)) {
        seen.set(name, mark);
      }
    }
  });

  return Array.from(seen.values());
};
