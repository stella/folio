import { panic } from "better-result";

import { assertXmlResourceLimits } from "@stll/docx-core/xml/limits";

import type {
  Shape,
  ShapeAlternateContent,
  ShapeContent,
} from "../types/content";
import { canonicalJson } from "../utils/canonicalJson";
import { OOXML_NAMESPACES } from "./serializer/partNamespaces";
import {
  captureVerbatimXml,
  isSafeCapturedXmlDocument,
} from "./verbatimCapture";
import {
  cloneWithXmlnsDeclarations,
  getChildElements,
  getLocalName,
  getNamespacePrefix,
  getNamespaceUri,
  NAMESPACES,
  parseXmlDocument,
  WORDPROCESSINGML_NAMESPACE_URIS,
  type XmlElement,
} from "./xmlParser";

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
  const selfContainedAlternateContent = cloneWithXmlnsDeclarations(
    alternateContent,
    {},
  );
  return {
    verbatimXml: captureVerbatimXml(selfContainedAlternateContent),
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

type DescendantsNamedOptions = {
  readonly root: XmlElement;
  readonly namespaceUris: ReadonlySet<string>;
  readonly localName: string;
};

const descendantsNamed = ({
  root,
  namespaceUris,
  localName,
}: DescendantsNamedOptions): XmlElement[] => {
  const result: XmlElement[] = [];
  const visit = (element: XmlElement): void => {
    for (const child of getChildElements(element)) {
      if (
        namespaceUris.has(getNamespaceUri(child) ?? "") &&
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

/**
 * The freshly written body leans on its document's root for prefixes such as
 * `w14`, which a captured alternate content does not declare for itself.
 */
const namespaceDeclarationsFor = (body: XmlElement): Record<string, string> => {
  const declarations: Record<string, string> = {};
  const visit = (element: XmlElement): void => {
    for (const name of [
      element.name ?? "",
      ...Object.keys(element.attributes ?? {}),
    ]) {
      const prefix = getNamespacePrefix(name);
      // The writer's own prefixes too: a tracked change carries folio's
      // review history attributes.
      const uri =
        prefix !== null && Object.hasOwn(OOXML_NAMESPACES, prefix)
          ? OOXML_NAMESPACES[prefix as keyof typeof OOXML_NAMESPACES].uri
          : undefined;
      if (prefix !== "xmlns" && uri !== undefined) {
        declarations[`xmlns:${prefix}`] = uri;
      }
    }
    for (const child of element.elements ?? []) visit(child);
  };
  visit(body);
  return declarations;
};

const updateAlternateContentText = (
  capturedXml: string,
  drawingXml: string,
): string | undefined => {
  assertXmlResourceLimits({ xml: capturedXml });
  assertXmlResourceLimits({ xml: drawingXml });
  const alternateContent = parseXmlDocument(capturedXml);
  const drawing = parseXmlDocument(
    `<root xmlns:w="${NAMESPACES.w}">${drawingXml}</root>`,
  );
  if (alternateContent === null || drawing === null) {
    panic(
      "Captured shape alternate content failed to parse during serialization.",
    );
  }
  const choice = branchNamed(alternateContent, "Choice");
  const fallback = branchNamed(alternateContent, "Fallback");
  const choiceBodies = choice
    ? descendantsNamed({
        root: choice,
        namespaceUris: WORDPROCESSINGML_NAMESPACE_URIS,
        localName: "txbxContent",
      })
    : [];
  const fallbackBodies = fallback
    ? descendantsNamed({
        root: fallback,
        namespaceUris: WORDPROCESSINGML_NAMESPACE_URIS,
        localName: "txbxContent",
      })
    : [];
  const drawingBodies = descendantsNamed({
    root: drawing,
    namespaceUris: WORDPROCESSINGML_NAMESPACE_URIS,
    localName: "txbxContent",
  });
  if (
    choiceBodies.length !== 1 ||
    fallbackBodies.length !== 1 ||
    drawingBodies.length !== 1
  ) {
    return undefined;
  }
  const choiceBody = choiceBodies.at(0);
  const fallbackBody = fallbackBodies.at(0);
  const drawingBody = drawingBodies.at(0);
  if (
    choiceBody === undefined ||
    fallbackBody === undefined ||
    drawingBody === undefined
  ) {
    return undefined;
  }
  const qualifiedName = drawingBody.name ?? "txbxContent";
  const prefix = getNamespacePrefix(qualifiedName);
  const declaration = prefix === null ? "xmlns" : `xmlns:${prefix}`;
  const declarations = namespaceDeclarationsFor(drawingBody);
  // A prefix a consumer may not know is ignorable where it is declared:
  // the captured element's ancestors in the part do not declare it.
  const ignorable = Object.keys(declarations)
    .map((name) => name.slice("xmlns:".length))
    .filter(
      (declared) =>
        Object.hasOwn(OOXML_NAMESPACES, declared) &&
        OOXML_NAMESPACES[declared as keyof typeof OOXML_NAMESPACES].ignorable,
    );
  for (const body of [choiceBody, fallbackBody]) {
    body.name = qualifiedName;
    const stated = body.attributes?.["mc:Ignorable"];
    body.attributes = {
      ...declarations,
      ...body.attributes,
      [declaration]: NAMESPACES.w,
      ...(ignorable.length === 0
        ? {}
        : {
            "xmlns:mc": NAMESPACES.mc,
            "mc:Ignorable": [
              ...new Set([
                ...(typeof stated === "string" ? stated.split(/\s+/u) : []),
                ...ignorable,
              ]),
            ]
              .filter((value) => value !== "")
              .join(" "),
          }),
    };
    body.elements = drawingBody.elements ?? [];
  }
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
 * Replay requires the modeled shape to match the capture. Structural edits
 * regenerate a bare drawing rather than replaying stale geometry.
 */
export const replayableShapeAlternateContent = (
  content: ShapeContent,
): string | undefined => {
  const captured = content.alternateContent;
  if (captured === undefined) {
    return undefined;
  }
  if (captured.verbatimFingerprint === shapeFingerprint(content.shape)) {
    return captured.verbatimXml;
  }
  return undefined;
};

/** Synchronize only a text-only edit, after the caller needs regenerated drawing markup. */
export const synchronizedTextShapeAlternateContent = (
  content: ShapeContent,
  regeneratedDrawingXml: string,
): string | undefined => {
  const captured = content.alternateContent;
  if (
    captured === undefined ||
    captured.verbatimStructureFingerprint === undefined ||
    captured.verbatimStructureFingerprint !==
      shapeStructureFingerprint(content.shape)
  ) {
    return undefined;
  }
  return updateAlternateContentText(
    captured.verbatimXml,
    regeneratedDrawingXml,
  );
};
