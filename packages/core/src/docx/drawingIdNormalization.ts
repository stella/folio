import { captureVerbatimXml } from "./verbatimCapture";
import { toTransitionalNamespaceUri } from "./transitionalSpelling";
import type { DrawingContent, Image, Shape } from "../types/document";
import { panic } from "better-result";
import { isNewDataUrlDrawing } from "./newImage";
import {
  visitDocxParagraphs,
  visitParagraphRuns,
  type DocxParagraphSurfaces,
} from "./paragraphTraversal";
import {
  getLocalName,
  getNamespaceUri,
  NAMESPACES,
  OOXML_NAMESPACE_SCOPE,
  parseXml,
} from "./xmlParser";
import type { XmlElement } from "./xmlParser";

const GENERATED_DRAWING_ID_START = 100_000;

type DrawingWithId = Image | Shape;

type DrawingIdEntry = {
  drawing: DrawingWithId;
  generateWhenMissing: boolean;
  rawDrawing?: DrawingContent;
};

const isDetachedRawDrawing = (drawing: DrawingContent | undefined): drawing is DrawingContent =>
  drawing?.rawXml !== undefined && isNewDataUrlDrawing(drawing);

const reassignRawDrawingId = ({ xml, id }: { xml: string; id: string }): string | null => {
  const root = parseXml(xml, OOXML_NAMESPACE_SCOPE);
  let foundDocPr = false;
  const visit = (element: XmlElement): void => {
    if (
      (toTransitionalNamespaceUri(getNamespaceUri(element) ?? "") === NAMESPACES.wp &&
        getLocalName(element.name) === "docPr") ||
      (toTransitionalNamespaceUri(getNamespaceUri(element) ?? "") === NAMESPACES.pic &&
        getLocalName(element.name) === "cNvPr")
    ) {
      for (const name of Object.keys(element.attributes ?? {})) {
        if (getLocalName(name) === "id" && element.attributes) {
          element.attributes[name] = id;
          if (
            toTransitionalNamespaceUri(getNamespaceUri(element) ?? "") === NAMESPACES.wp &&
            getLocalName(element.name) === "docPr"
          )
            foundDocPr = true;
        }
      }
    }
    for (const child of element.elements ?? []) visit(child);
  };
  visit(root);
  return foundDocPr ? (root.elements ?? []).map(captureVerbatimXml).join("") : null;
};

const needsGeneratedId = ({ id }: DrawingWithId): boolean =>
  id === undefined || id === "" || id === "0";

export const normalizeDrawingIds = (surfaces: DocxParagraphSurfaces): void => {
  const entries: DrawingIdEntry[] = [];

  visitDocxParagraphs(surfaces, (paragraph) => {
    visitParagraphRuns(paragraph, (run) => {
      for (const content of run.content) {
        if (content.type === "shape") {
          entries.push({ drawing: content.shape, generateWhenMissing: true });
          continue;
        }
        if (content.type === "drawing") {
          entries.push({
            drawing: content.image,
            generateWhenMissing: content.rawXml === undefined,
            ...(content.rawXml !== undefined ? { rawDrawing: content } : {}),
          });
        }
      }
    });
  });

  const usedIds = new Set(
    entries.flatMap(({ drawing, rawDrawing }) =>
      needsGeneratedId(drawing) || isDetachedRawDrawing(rawDrawing) ? [] : [drawing.id],
    ),
  );
  let nextId = GENERATED_DRAWING_ID_START;

  for (const { drawing, generateWhenMissing } of entries) {
    if (!generateWhenMissing || !needsGeneratedId(drawing)) {
      continue;
    }
    while (usedIds.has(String(nextId))) {
      nextId += 1;
    }
    drawing.id = String(nextId);
    usedIds.add(drawing.id);
    nextId += 1;
  }

  for (const { drawing, rawDrawing } of entries) {
    if (!isDetachedRawDrawing(rawDrawing)) {
      continue;
    }
    if (!needsGeneratedId(drawing) && !usedIds.has(drawing.id)) {
      usedIds.add(drawing.id);
      continue;
    }
    while (usedIds.has(String(nextId))) {
      nextId += 1;
    }
    const sourceXml = rawDrawing.rawXml;
    if (sourceXml === undefined) {
      panic("Detached raw drawing must retain its XML.");
    }
    const rawXml = reassignRawDrawingId({ xml: sourceXml, id: String(nextId) });
    if (rawXml === null) {
      panic("Detached raw drawing must contain wp:docPr.");
    }
    drawing.id = String(nextId);
    rawDrawing.rawXml = rawXml;
    usedIds.add(drawing.id);
    nextId += 1;
  }
};
