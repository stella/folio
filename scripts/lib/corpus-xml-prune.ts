/**
 * Remove XML subtrees from an OOXML part while keeping everything else intact.
 *
 * Elements are addressed by their position in a pre-order walk of the original
 * part, so an address means the same thing however many other elements were
 * dropped. Dropping an element drops its subtree, and the walk still consumes
 * the addresses inside it so the surviving addresses do not shift.
 *
 * The parse is lossy in the ways `fast-xml-parser` is lossy (entity spelling,
 * attribute order, insignificant whitespace), which is why the minimiser probes
 * a no-op rebuild first and refuses element-level work on a part whose failure
 * the rebuild alone changes.
 */

import { XMLBuilder, XMLParser } from "fast-xml-parser";

const ATTRIBUTE_KEY = ":@";
const ATTRIBUTE_PREFIX = "@_";

type OrderedNode = Record<string, unknown>;

const parserOptions = {
  preserveOrder: true,
  ignoreAttributes: false,
  attributeNamePrefix: ATTRIBUTE_PREFIX,
  parseTagValue: false,
  parseAttributeValue: false,
  trimValues: false,
} as const;

/** The element name a preserve-order node carries, or null for text and processing nodes. */
const elementTag = (node: OrderedNode): string | null => {
  for (const key of Object.keys(node)) {
    if (key !== ATTRIBUTE_KEY && !key.startsWith("#") && !key.startsWith("?")) {
      return key;
    }
  }
  return null;
};

export type PrunableXml = {
  /** Every removable element address, in pre-order. */
  addresses: number[];
  render: (keep: ReadonlySet<number>) => string;
};

type PruneState = { next: number; keep: (address: number) => boolean };

const pruneNodes = (nodes: readonly OrderedNode[], state: PruneState): OrderedNode[] => {
  const kept: OrderedNode[] = [];
  for (const node of nodes) {
    const tag = elementTag(node);
    if (tag === null) {
      kept.push(node);
      continue;
    }
    const address = state.next;
    state.next += 1;
    const children = pruneNodes((node[tag] ?? []) as OrderedNode[], state);
    if (state.keep(address)) {
      kept.push({ ...node, [tag]: children });
    }
  }
  return kept;
};

export const prepareXmlPruning = (xml: string): PrunableXml => {
  const parsed = new XMLParser(parserOptions).parse(xml) as OrderedNode[];
  const counter: PruneState = { next: 0, keep: () => true };
  pruneNodes(parsed, counter);
  const builder = new XMLBuilder({ ...parserOptions, suppressEmptyNode: false });
  return {
    addresses: Array.from({ length: counter.next }, (_, index) => index),
    render: (keep) =>
      builder.build(
        pruneNodes(parsed, { next: 0, keep: (address) => keep.has(address) }),
      ) as string,
  };
};
