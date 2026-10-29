/**
 * Remove XML subtrees from an OOXML part while keeping everything else intact.
 *
 * Elements are addressed by their position in a pre-order walk of the original
 * part, so an address means the same thing however many other elements were
 * dropped. Dropping an element drops its subtree, and the walk still consumes
 * the addresses inside it so the surviving addresses do not shift.
 *
 * The rebuild is lossy where a parse is (entity spelling, quoting, comments,
 * insignificant whitespace inside tags), which is why the minimiser probes a
 * no-op rebuild first and refuses element-level work on a part whose failure
 * the rebuild alone changes.
 */

import { escapeXmlAttribute, escapeXmlText } from "@stll/docx-core";
import { readXml, type XmlElement } from "@stll/docx-core/xml";

export type PrunableXml = {
  /** Every removable element address, in pre-order. */
  addresses: number[];
  render: (keep: ReadonlySet<number>) => string;
};

type PruneState = { next: number; keep: (address: number) => boolean };

/** Write `nodes`, dropping each element whose pre-order address `keep` refuses. */
const writeNodes = (
  nodes: readonly XmlElement[],
  state: PruneState,
  out: string[],
): void => {
  for (const node of nodes) {
    if (node.type === "text") {
      out.push(escapeXmlText(String(node.text ?? "")));
      continue;
    }
    if (node.type === "instruction") {
      const body = node.instruction ? ` ${node.instruction}` : "";
      out.push(`<?${node.name ?? ""}${body}?>`);
      continue;
    }
    const address = state.next;
    state.next += 1;
    const kept = state.keep(address);
    const start = out.length;
    if (kept) {
      out.push(`<${node.name ?? ""}`);
      for (const [name, value] of Object.entries(node.attributes ?? {})) {
        if (value !== undefined)
          out.push(` ${name}="${escapeXmlAttribute(String(value))}"`);
      }
      out.push(">");
    }
    // The subtree is walked even when dropped, so its addresses are consumed.
    writeNodes(node.elements ?? [], state, out);
    if (kept) {
      out.push(`</${node.name ?? ""}>`);
    } else {
      out.length = start;
    }
  }
};

export const prepareXmlPruning = (xml: string): PrunableXml => {
  const read = readXml(xml, {
    conformance: "tolerant",
    instructions: "retain",
  });
  if (read.status === "rejected") {
    throw new Error(`XML part could not be read: ${read.reason}`);
  }
  const nodes = read.value.elements ?? [];
  const render = (keep: (address: number) => boolean): string => {
    const out: string[] = [];
    writeNodes(nodes, { next: 0, keep }, out);
    return out.join("");
  };
  const counter: PruneState = { next: 0, keep: () => true };
  writeNodes(nodes, counter, []);
  return {
    addresses: Array.from({ length: counter.next }, (_, index) => index),
    render: (keep) => render((address) => keep.has(address)),
  };
};
