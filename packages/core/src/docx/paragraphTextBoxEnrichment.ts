import type {
  MediaFile,
  Paragraph,
  RelationshipMap,
  Run,
  Shape,
  ShapeContent,
  Theme,
} from "../types/document";
import type { NumberingMap } from "./numberingParser";
import { parseParagraph } from "./paragraphParser";
import type { StyleMap } from "./styleParser";
import {
  getTextBoxContentElement,
  isTextBoxDrawing,
  parseTextBox,
  parseTextBoxContent,
} from "./textBoxParser";
import { findDeep, getChildElements, getLocalName, type XmlElement } from "./xmlParser";

export const enrichParagraphTextBoxes = (
  paragraph: Paragraph,
  paraXml: XmlElement,
  styles: StyleMap | null,
  theme: Theme | null,
  numbering: NumberingMap | null,
  rels: RelationshipMap | null,
  media: Map<string, MediaFile> | null,
): void => {
  const xmlChildren = getChildElements(paraXml);
  let parsedIndex = 0;
  let lastConsumedRun: Run | undefined;

  for (const xmlChild of xmlChildren) {
    if (getLocalName(xmlChild.name ?? "") !== "r") {
      if (
        parsedIndex < paragraph.content.length &&
        paragraph.content[parsedIndex]?.type !== "run"
      ) {
        parsedIndex += 1;
      }
      continue;
    }

    const { textBoxDrawings, hasNonTextBoxContent } = scanRunForTextBoxDrawings(xmlChild);

    const parsedContent = paragraph.content[parsedIndex];
    const parsedRun: Run | undefined = parsedContent?.type === "run" ? parsedContent : undefined;
    const targetRun = parsedRun ?? (hasNonTextBoxContent ? lastConsumedRun : undefined);

    for (const runEl of textBoxDrawings) {
      const textBox = parseTextBox(runEl);
      if (!textBox) {
        continue;
      }

      const wsp = findDeep(runEl, "wps", "wsp");
      if (wsp) {
        const txbxContentEl = getTextBoxContentElement(wsp);
        if (txbxContentEl) {
          textBox.content = parseTextBoxContent(
            txbxContentEl,
            parseParagraph,
            null,
            styles,
            theme,
            numbering,
            rels ?? undefined,
            media ?? undefined,
          );
        }
      }

      const shape: Shape = {
        type: "shape",
        shapeType: "textBox",
        size: textBox.size,
        ...(textBox.position !== undefined ? { position: textBox.position } : {}),
        ...(textBox.wrap !== undefined ? { wrap: textBox.wrap } : {}),
        ...(textBox.fill !== undefined ? { fill: textBox.fill } : {}),
        ...(textBox.outline !== undefined ? { outline: textBox.outline } : {}),
        textBody: {
          content: textBox.content,
          ...(textBox.autoFit !== undefined ? { autoFit: textBox.autoFit } : {}),
          ...(textBox.margins !== undefined ? { margins: textBox.margins } : {}),
        },
      };
      if (textBox.id) {
        shape.id = textBox.id;
      }

      const shapeContent: ShapeContent = { type: "shape", shape };

      if (targetRun && hasNonTextBoxContent) {
        targetRun.content.push(shapeContent);
      } else {
        const newRun: Run = { type: "run", content: [shapeContent] };
        paragraph.content.splice(parsedIndex, 0, newRun);
        lastConsumedRun = newRun;
        parsedIndex += 1;
      }
    }

    if (hasNonTextBoxContent && parsedRun) {
      lastConsumedRun = parsedRun;
      parsedIndex += 1;
    }
  }
};

type TextBoxRunScan = {
  textBoxDrawings: XmlElement[];
  hasNonTextBoxContent: boolean;
};

const scanRunForTextBoxDrawings = (xmlRun: XmlElement): TextBoxRunScan => {
  const textBoxDrawings: XmlElement[] = [];
  let hasNonTextBoxContent = false;

  const visitDrawing = (drawingEl: XmlElement): void => {
    if (isTextBoxDrawing(drawingEl)) {
      textBoxDrawings.push(drawingEl);
      return;
    }
    hasNonTextBoxContent = true;
  };

  for (const el of getChildElements(xmlRun)) {
    const name = getLocalName(el.name ?? "");
    if (name === "rPr") {
      continue;
    }
    if (name === "drawing") {
      visitDrawing(el);
      continue;
    }
    if (name === "AlternateContent") {
      const branches = getChildElements(el);
      const choice = branches.find((branch) => getLocalName(branch.name ?? "") === "Choice");
      const fallback = branches.find((branch) => getLocalName(branch.name ?? "") === "Fallback");
      const tryBranch = (branch: XmlElement | undefined): boolean => {
        if (!branch) {
          return false;
        }
        let found = false;
        for (const innerEl of getChildElements(branch)) {
          if (getLocalName(innerEl.name ?? "") === "drawing") {
            visitDrawing(innerEl);
            found = true;
          }
        }
        return found;
      };
      let foundInBranch = tryBranch(choice);
      if (!foundInBranch) {
        foundInBranch = tryBranch(fallback);
      }
      if (!foundInBranch) {
        hasNonTextBoxContent = true;
      }
      continue;
    }
    hasNonTextBoxContent = true;
  }

  return { textBoxDrawings, hasNonTextBoxContent };
};
