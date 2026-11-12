import {
  parseSplitResolutionPiece,
  type SplitResolutionPiece,
} from "./reviewResolutionProvenance";
import type { XmlElement } from "./xmlParser";

/**
 * The pieces of changes and wrappers the serializer wrote as several
 * elements around a lifted link, by the record each was read into. The
 * paragraph parser puts them back together once a paragraph is read.
 */
export const splitPieces: WeakMap<object, SplitResolutionPiece> = new WeakMap();

/** Record that `record` was read from `node`, a piece of a split element. */
export const tagSplitPiece = (node: XmlElement, record: object): void => {
  const split = parseSplitResolutionPiece(node);
  if (split !== undefined) splitPieces.set(record, split);
};
