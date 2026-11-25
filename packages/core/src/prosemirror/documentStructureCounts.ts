import type { Node as PMNode } from "prosemirror-model";

import type { SectionProperties } from "../types/document";
import { canonicalJson } from "../utils/canonicalJson";
import { sectionPropertiesOf } from "./sectionCarrier";

export type DocumentStructureCounts = {
  blockStructureFingerprint: string;
  sectionRecords: ReadonlySet<SectionProperties>;
  paragraphs: number;
  /**
   * Distinct section records the paragraphs carry, which is the number of
   * `w:sectPr` the package states for them. A split copies a node's attrs, so
   * both halves of a split section-ending paragraph hold one record until the
   * from-leg keeps it on the trailing half: one section, not two.
   */
  sectionEndpoints: number;
  sectionEndpointFingerprint: string;
};

type SectionEndpointRecord = {
  path: string;
  paraId: unknown;
  sectionProperties: SectionProperties;
};

/**
 * What one block contributes to {@link countDocumentStructure}, with paths
 * relative to the block (`""` is the block itself). A node is immutable, so
 * an unchanged block's contribution is computed once.
 */
type BlockStructure = {
  /** The block's structure records, as JSON. */
  records: string;
  paragraphs: number;
  endpoints: readonly SectionEndpointRecord[];
};

const blockStructures = new WeakMap<PMNode, BlockStructure>();

const relativePath = (parentPath: string, index: number): string =>
  parentPath.length === 0 ? `${index}` : `${parentPath}.${index}`;

const blockStructureOf = (block: PMNode): BlockStructure => {
  const cached = blockStructures.get(block);
  if (cached) return cached;
  const records: unknown[] = [];
  const endpoints: SectionEndpointRecord[] = [];
  let paragraphs = 0;
  const visit = (node: PMNode, path: string): void => {
    records.push(
      node.type.name === "paragraph"
        ? [path, node.type.name, node.attrs["paraId"] ?? null]
        : [path, node.type.name, canonicalJson(node.attrs)],
    );
    if (node.type.name === "paragraph") {
      paragraphs++;
      const sectionProperties = sectionPropertiesOf(node);
      // The record is the whole endpoint: its `sectionStart` is inside the
      // canonical JSON of the endpoints, so a type change is a fingerprint
      // change.
      if (sectionProperties !== null)
        endpoints.push({
          path,
          paraId: node.attrs["paraId"] ?? null,
          sectionProperties,
        });
      return;
    }
    node.forEach((child, _offset, index) =>
      visit(child, relativePath(path, index)),
    );
  };
  visit(block, "");
  const structure = {
    records: JSON.stringify(records),
    paragraphs,
    endpoints,
  };
  blockStructures.set(block, structure);
  return structure;
};

/**
 * Count the structural values the tracker compares for every transaction.
 *
 * Each top-level block's records are kept for the block node, so a keystroke
 * recomputes only the block it changed. The fingerprint joins the blocks'
 * records in order: two documents' fingerprints are equal exactly when their
 * full record lists (every block and nested block with its path) are.
 */
export function countDocumentStructure(doc: PMNode): DocumentStructureCounts {
  let paragraphs = 0;
  const sectionRecords = new Set<SectionProperties>();
  const blockRecords: string[] = [];
  const endpointRecords: SectionEndpointRecord[] = [];
  doc.forEach((block, _offset, index) => {
    const structure = blockStructureOf(block);
    blockRecords.push(structure.records);
    paragraphs += structure.paragraphs;
    for (const endpoint of structure.endpoints) {
      sectionRecords.add(endpoint.sectionProperties);
      endpointRecords.push({
        ...endpoint,
        path:
          endpoint.path.length === 0 ? `${index}` : `${index}.${endpoint.path}`,
      });
    }
  });
  return {
    // JSON escapes control characters, so the separator never occurs inside
    // a block's records.
    blockStructureFingerprint: blockRecords.join("\u0001"),
    sectionRecords,
    paragraphs,
    sectionEndpoints: sectionRecords.size,
    sectionEndpointFingerprint: canonicalJson(endpointRecords),
  };
}
