import type { Shape, ShapeAlternateContent, ShapeContent } from "../types/content";
import { canonicalJson } from "../utils/canonicalJson";
import { captureVerbatimXml } from "./verbatimCapture";
import { getChildElements, getLocalName, type XmlElement } from "./xmlParser";

/**
 * Every modeled field, text body included: the Fallback repeats the shape's
 * geometry and its text, so an edit to either makes the captured element stale.
 */
const shapeFingerprint = (shape: Shape): string => canonicalJson(shape);

/**
 * Whether `element` is the only child element of its `mc:Choice` or
 * `mc:Fallback` branch.
 *
 * A branch holding more than the one drawing cannot be written back by the
 * drawing's shape alone: the siblings are parsed into their own run content,
 * and replaying the whole element would write them twice.
 */
const isSoleBranchChild = (branch: XmlElement, element: XmlElement): boolean => {
  const children = getChildElements(branch);
  return children.length === 1 && children[0] === element;
};

/**
 * Capture the `mc:AlternateContent` a shape was read from, when the shape's
 * drawing is the sole child of the branch it came from.
 */
export const captureShapeAlternateContent = ({
  shape,
  alternateContent,
  branch,
  drawing,
}: {
  shape: Shape;
  alternateContent: XmlElement;
  branch: XmlElement;
  drawing: XmlElement;
}): ShapeAlternateContent | undefined => {
  if (getLocalName(alternateContent.name ?? "") !== "AlternateContent") {
    return undefined;
  }
  if (!isSoleBranchChild(branch, drawing)) {
    return undefined;
  }
  return {
    verbatimXml: captureVerbatimXml(alternateContent),
    verbatimFingerprint: shapeFingerprint(shape),
  };
};

/** Re-fingerprint a capture for a shape rebuilt from an unedited projection. */
export const refingerprintShapeAlternateContent = (
  shape: Shape,
  verbatimXml: string,
): ShapeAlternateContent => ({ verbatimXml, verbatimFingerprint: shapeFingerprint(shape) });

/**
 * The captured `mc:AlternateContent` to write in place of the shape's drawing,
 * or undefined when the shape must be regenerated.
 *
 * Replay requires the modeled shape to still match the capture. An edited shape
 * is regenerated as a bare `w:drawing` without the stale Fallback.
 */
export const replayableShapeAlternateContent = (content: ShapeContent): string | undefined => {
  const captured = content.alternateContent;
  if (captured === undefined) {
    return undefined;
  }
  return captured.verbatimFingerprint === shapeFingerprint(content.shape)
    ? captured.verbatimXml
    : undefined;
};
