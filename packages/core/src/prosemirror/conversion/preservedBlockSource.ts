import type { Node as PMNode } from "prosemirror-model";
import type { PreservedBlock } from "../../types/document";

const projectedSources = new WeakMap<PMNode, PreservedBlock>();
const extractedSources = new WeakMap<PreservedBlock, PreservedBlock>();

/** Called only while projecting a tracked main story, including its nested blocks. */
export const rememberPreservedBlockProjection = (node: PMNode, source: PreservedBlock): void => {
  projectedSources.set(node, source);
};

/** Extraction carries a candidate reference; the base model must still own it. */
export const inheritPreservedBlockProjection = (target: PreservedBlock, node: PMNode): void => {
  const source = projectedSources.get(node);
  if (source !== undefined) extractedSources.set(target, source);
};

export const getPreservedBlockProjectionSource = (
  block: PreservedBlock,
): PreservedBlock | undefined => extractedSources.get(block);
