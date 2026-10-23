import { panic } from "better-result";

import type {
  Shape,
  ShapeAlternateContent,
  ShapeContent,
} from "../types/content";
import { canonicalJson } from "../utils/canonicalJson";
import {
  captureVerbatimXml,
  isSafeCapturedXmlDocument,
} from "./verbatimCapture";
import {
  getChildElements,
  getLocalName,
  getNamespacePrefix,
  getNamespaceUri,
  NAMESPACES,
  parseXmlDocument,
  type XmlElement,
} from "./xmlParser";
import { assertXmlResourceLimits } from "./xmlResourceLimits";

/**
 * Every modeled field, text body included: the Fallback repeats the shape's
 * geometry and its text, so an edit to either makes the captured element stale.
 */
const shapeFingerprint = (shape: Shape): string => canonicalJson(shape);

const shapeStructureFingerprint = (shape: Shape): string => {
  const { textBody, ...shapeProperties } = shape;
  if (textBody === undefined) {
    return canonicalJson(shapeProperties);
  }
  const { content: _content, ...textBodyProperties } = textBody;
  return canonicalJson({ ...shapeProperties, textBody: textBodyProperties });
};

/**
 * Whether `element` is the only child element of its `mc:Choice` or
 * `mc:Fallback` branch.
 *
 * A branch holding more than the one drawing cannot be written back by the
 * drawing's shape alone: the siblings are parsed into their own run content,
 * and replaying the whole element would write them twice.
 */
const isSoleBranchChild = (
  branch: XmlElement,
  element: XmlElement,
): boolean => {
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
    verbatimStructureFingerprint: shapeStructureFingerprint(shape),
  };
};

/** Re-fingerprint a capture for a shape rebuilt from an unedited projection. */
export const refingerprintShapeAlternateContent = (
  shape: Shape,
  verbatimXml: string,
): ShapeAlternateContent => ({
  verbatimXml,
  verbatimFingerprint: shapeFingerprint(shape),
  verbatimStructureFingerprint: shapeStructureFingerprint(shape),
});

const descendantsNamed = (
  root: XmlElement,
  namespaceUri: string,
  localName: string,
): XmlElement[] => {
  const result: XmlElement[] = [];
  const visit = (element: XmlElement): void => {
    for (const child of getChildElements(element)) {
      if (
        getNamespaceUri(child) === namespaceUri &&
        getLocalName(child.name ?? "") === localName
      ) {
        result.push(child);
      }
      visit(child);
    }
  };
  visit(root);
  return result;
};

const branchNamed = (
  root: XmlElement,
  localName: string,
): XmlElement | undefined =>
  getChildElements(root).find(
    (child) =>
      getNamespaceUri(child) === NAMESPACES.mc &&
      getLocalName(child.name ?? "") === localName,
  );

const updateFallbackText = (
  capturedXml: string,
  drawingXml: string,
): string | undefined => {
  assertXmlResourceLimits({ xml: capturedXml });
  assertXmlResourceLimits({ xml: drawingXml });
  const alternateContent = parseXmlDocument(capturedXml);
  const drawing = parseXmlDocument(drawingXml);
  if (alternateContent === null || drawing === null) {
    panic(
      "Captured shape alternate content failed to parse during serialization.",
    );
  }
  const fallback = branchNamed(alternateContent, "Fallback");
  const fallbackBodies = fallback
    ? descendantsNamed(fallback, NAMESPACES.w, "txbxContent")
    : [];
  const drawingBodies = descendantsNamed(drawing, NAMESPACES.w, "txbxContent");
  if (fallbackBodies.length !== 1 || drawingBodies.length !== 1) {
    return undefined;
  }
  const fallbackBody = fallbackBodies[0];
  const drawingBody = drawingBodies[0];
  const prefix = getNamespacePrefix(drawingBody.name ?? "");
  const declaration = prefix === null ? "xmlns" : `xmlns:${prefix}`;
  fallbackBody.attributes = {
    ...fallbackBody.attributes,
    [declaration]: NAMESPACES.w,
  };
  fallbackBody.elements = drawingBody.elements;
  const synchronizedXml = captureVerbatimXml(alternateContent);
  if (!isSafeCapturedXmlDocument(synchronizedXml)) {
    panic("Synchronized shape alternate content failed XML validation.");
  }
  return synchronizedXml;
};

/**
 * The captured `mc:AlternateContent` to write in place of the shape's drawing,
 * or undefined when the shape must be regenerated.
 *
 * Replay requires the modeled shape to match the capture. Text-only edits
 * regenerate the drawing and synchronize its VML Fallback; structural edits
 * regenerate a bare drawing rather than replaying stale geometry.
 */
export const replayableShapeAlternateContent = (
  content: ShapeContent,
  regeneratedDrawingXml?: string,
): string | undefined => {
  const captured = content.alternateContent;
  if (captured === undefined) {
    return undefined;
  }
  if (captured.verbatimFingerprint === shapeFingerprint(content.shape)) {
    return captured.verbatimXml;
  }
  if (
    regeneratedDrawingXml !== undefined &&
    captured.verbatimStructureFingerprint !== undefined &&
    captured.verbatimStructureFingerprint ===
      shapeStructureFingerprint(content.shape)
  ) {
    return updateFallbackText(captured.verbatimXml, regeneratedDrawingXml);
  }
  return undefined;
};
