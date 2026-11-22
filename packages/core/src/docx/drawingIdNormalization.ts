// PARSE-WARNING-EXEMPT: renumbers drawing ids to keep them unique; an id is
// not content, and every reference is remapped with it.

import { panic } from "better-result";

import {
  assertValidOoxmlNumericId,
  isValidOoxmlNumericId,
} from "@stll/docx-core";

import type { DrawingContent, Image, Shape } from "../types/document";
import type { DeepReadonly } from "../types/documentModel";
import { isNewDataUrlDrawing } from "./newImage";
import {
  type DocxParagraphSurfaces,
  visitDocxParagraphs,
  visitParagraphRuns,
} from "./paragraphTraversal";
import { toTransitionalNamespaceUri } from "./transitionalSpelling";
import { captureVerbatimXml } from "./verbatimCapture";
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

const isDetachedRawDrawing = (
  drawing: DrawingContent | undefined,
): drawing is DrawingContent =>
  drawing?.rawXml !== undefined && isNewDataUrlDrawing(drawing);

const reassignRawDrawingId = ({
  xml,
  id,
}: {
  xml: string;
  id: string;
}): string | null => {
  const root = parseXml(xml, OOXML_NAMESPACE_SCOPE);
  let foundDocPr = false;
  const visit = (element: XmlElement): void => {
    if (
      (toTransitionalNamespaceUri(getNamespaceUri(element) ?? "") ===
        NAMESPACES.wp &&
        getLocalName(element.name) === "docPr") ||
      (toTransitionalNamespaceUri(getNamespaceUri(element) ?? "") ===
        NAMESPACES.pic &&
        getLocalName(element.name) === "cNvPr")
    ) {
      for (const name of Object.keys(element.attributes ?? {})) {
        if (getLocalName(name) === "id" && element.attributes) {
          element.attributes[name] = id;
          if (
            toTransitionalNamespaceUri(getNamespaceUri(element) ?? "") ===
              NAMESPACES.wp &&
            getLocalName(element.name) === "docPr"
          )
            foundDocPr = true;
        }
      }
    }
    for (const child of element.elements ?? []) visit(child);
  };
  visit(root);
  return foundDocPr
    ? (root.elements ?? []).map(captureVerbatimXml).join("")
    : null;
};

/** Drawing IDs are package-local identifiers, not image content. */
export const drawingXmlWithoutIdentity = (xml: string): string =>
  reassignRawDrawingId({ xml, id: "0" }) ?? xml;

const needsGeneratedId = ({ id }: DrawingWithId): boolean =>
  id === undefined || !isValidOoxmlNumericId(id, "unsigned32");

const RAW_DRAWING_ID = /<(?:[\w.-]+:)?docPr\s(?:[^>]*?\s)?id="(?<id>\d+)"/gu;

/** Every drawing id the surfaces use, in the model or in preserved markup. */
export const usedDrawingIds = (
  surfaces: DocxParagraphSurfaces,
): Set<number> => {
  const used = new Set<number>();
  const add = ({ id }: DrawingWithId) => {
    if (id !== undefined && isValidOoxmlNumericId(id, "unsigned32"))
      used.add(Number(id));
  };
  visitDocxParagraphs(surfaces, (paragraph) => {
    visitParagraphRuns(paragraph, (run) => {
      for (const content of run.content) {
        if (content.type === "shape") add(content.shape);
        if (content.type !== "drawing") continue;
        add(content.image);
        for (const match of content.rawXml?.matchAll(RAW_DRAWING_ID) ?? [])
          used.add(Number(match.groups?.["id"]));
      }
    });
  });
  return used;
};

/** A drawing id `used` does not hold yet, added to it. */
export const takeDrawingId = (used: Set<number>): string => {
  let id = GENERATED_DRAWING_ID_START;
  while (used.has(id)) id += 1;
  assertValidOoxmlNumericId({
    value: id,
    domain: "unsigned32",
    partPath: "drawing normalization",
    elementName: "docPr",
    attributeName: "id",
  });
  used.add(id);
  return String(id);
};

/** The drawing numbered `id`, in the model and in its preserved markup. */
export const withDrawingId = (
  drawing: DeepReadonly<DrawingContent>,
  id: string,
): DeepReadonly<DrawingContent> => ({
  ...drawing,
  image: { ...drawing.image, id },
  ...(drawing.rawXml === undefined
    ? {}
    : {
        rawXml:
          reassignRawDrawingId({ xml: drawing.rawXml, id }) ?? drawing.rawXml,
      }),
});

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
      needsGeneratedId(drawing) || isDetachedRawDrawing(rawDrawing)
        ? []
        : [Number(drawing.id)],
    ),
  );
  let nextId = GENERATED_DRAWING_ID_START;

  for (const { drawing, generateWhenMissing } of entries) {
    if (!generateWhenMissing || !needsGeneratedId(drawing)) {
      continue;
    }
    while (usedIds.has(nextId)) {
      nextId += 1;
    }
    assertValidOoxmlNumericId({
      value: nextId,
      domain: "unsigned32",
      partPath: "drawing normalization",
      elementName: "docPr",
      attributeName: "id",
    });
    drawing.id = String(nextId);
    usedIds.add(nextId);
    nextId += 1;
  }

  for (const { drawing, rawDrawing } of entries) {
    if (!isDetachedRawDrawing(rawDrawing)) {
      continue;
    }
    if (!needsGeneratedId(drawing) && !usedIds.has(Number(drawing.id))) {
      usedIds.add(Number(drawing.id));
      continue;
    }
    while (usedIds.has(nextId)) {
      nextId += 1;
    }
    assertValidOoxmlNumericId({
      value: nextId,
      domain: "unsigned32",
      partPath: "drawing normalization",
      elementName: "docPr",
      attributeName: "id",
    });
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
    usedIds.add(nextId);
    nextId += 1;
  }
};
