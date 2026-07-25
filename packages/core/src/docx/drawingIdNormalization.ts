import type { Image, Shape } from "../types/document";
import {
  visitDocxParagraphs,
  visitParagraphRuns,
  type DocxParagraphSurfaces,
} from "./paragraphTraversal";

const GENERATED_DRAWING_ID_START = 100_000;

type DrawingWithId = Image | Shape;

type DrawingIdEntry = {
  drawing: DrawingWithId;
  generateWhenMissing: boolean;
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
          });
        }
      }
    });
  });

  const usedIds = new Set(
    entries.flatMap(({ drawing }) => (needsGeneratedId(drawing) ? [] : [drawing.id])),
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
};
