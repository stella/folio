import type {
  ImagePosition,
  InlineSdt,
  MediaFile,
  Paragraph,
  ParagraphContent,
  RelationshipMap,
  Run,
  Shape,
  ShapeContent,
  Theme,
  TrackedRunChange,
} from "../types/document";
import { pixelsToEmu } from "../utils/units";
import type { NumberingMap } from "./numberingParser";
import { parseParagraph } from "./paragraphParser";
import type { StyleMap } from "./styleParser";
import {
  getTextBoxContentElement,
  isTextBoxDrawing,
  parseTextBox,
  parseTextBoxContent,
} from "./textBoxParser";
import type { TableParserFn } from "./textBoxParser";
import {
  findDeep,
  getAttribute,
  getChildElements,
  getLocalName,
  type XmlElement,
} from "./xmlParser";

const VML_HORIZONTAL_RELATIVES = new Set<ImagePosition["horizontal"]["relativeTo"]>([
  "character",
  "column",
  "insideMargin",
  "leftMargin",
  "margin",
  "outsideMargin",
  "page",
  "rightMargin",
]);
const VML_VERTICAL_RELATIVES = new Set<ImagePosition["vertical"]["relativeTo"]>([
  "insideMargin",
  "line",
  "margin",
  "outsideMargin",
  "page",
  "paragraph",
  "topMargin",
  "bottomMargin",
]);

const parseVmlStyle = (value: string | null): Record<string, string> => {
  const declarations: Record<string, string> = {};
  for (const declaration of value?.split(";") ?? []) {
    const separator = declaration.indexOf(":");
    if (separator < 0) {
      continue;
    }
    const key = declaration.slice(0, separator).trim().toLowerCase();
    if (key) {
      declarations[key] = declaration.slice(separator + 1).trim();
    }
  }
  return declarations;
};

const vmlLengthToPixels = (value: string | undefined): number | undefined => {
  const match = /^(?<amount>-?(?:\d+(?:\.\d+)?|\.\d+))\s*(?<unit>pt|in|px|cm|mm|pc)?$/iu.exec(
    value?.trim() ?? "",
  );
  const amount = Number.parseFloat(match?.groups?.["amount"] ?? "");
  if (!Number.isFinite(amount)) {
    return undefined;
  }
  switch (match?.groups?.["unit"]?.toLowerCase()) {
    case "pt":
      return (amount / 72) * 96;
    case "in":
      return amount * 96;
    case "cm":
      return (amount / 2.54) * 96;
    case "mm":
      return (amount / 25.4) * 96;
    case "pc":
      return amount * 16;
    case "px":
    case undefined:
      return amount;
    default:
      return undefined;
  }
};

const horizontalRelativeTo = (
  value: string | undefined,
): ImagePosition["horizontal"]["relativeTo"] => {
  for (const relative of VML_HORIZONTAL_RELATIVES) {
    if (relative.toLowerCase() === value?.toLowerCase()) {
      return relative;
    }
  }
  return "character";
};

const verticalRelativeTo = (value: string | undefined): ImagePosition["vertical"]["relativeTo"] => {
  for (const relative of VML_VERTICAL_RELATIVES) {
    if (relative.toLowerCase() === value?.toLowerCase()) {
      return relative;
    }
  }
  return "paragraph";
};

const parseVmlInsets = (
  textBoxEl: XmlElement,
): { left: number; top: number; right: number; bottom: number } | undefined => {
  const [left, top, right, bottom, extra] =
    getAttribute(textBoxEl, null, "inset")
      ?.split(",")
      .map((value) => vmlLengthToPixels(value)) ?? [];
  if (
    extra !== undefined ||
    left === undefined ||
    top === undefined ||
    right === undefined ||
    bottom === undefined
  ) {
    return undefined;
  }
  return {
    left: pixelsToEmu(left),
    top: pixelsToEmu(top),
    right: pixelsToEmu(right),
    bottom: pixelsToEmu(bottom),
  };
};

const parseVmlFill = (shapeEl: XmlElement): Shape["fill"] =>
  getAttribute(shapeEl, null, "filled")?.toLowerCase() === "f" ? { type: "none" } : undefined;

const vmlWrapType = (positioned: boolean, zIndex: number): NonNullable<Shape["wrap"]>["type"] => {
  if (!positioned) {
    return "inline";
  }
  if (Number.isFinite(zIndex) && zIndex < 0) {
    return "behind";
  }
  return "inFront";
};

export const enrichParagraphTextBoxes = (
  paragraph: Paragraph,
  paraXml: XmlElement,
  styles: StyleMap | null,
  theme: Theme | null,
  numbering: NumberingMap | null,
  rels: RelationshipMap | null,
  media: Map<string, MediaFile> | null,
  parseTable: TableParserFn,
): void => {
  enrichTextBoxRuns({
    content: paragraph.content,
    xmlChildren: getChildElements(paraXml),
    styles,
    theme,
    numbering,
    rels,
    media,
    parseTable,
  });
};

type EnrichTextBoxRunsParams = {
  content: ParagraphContent[];
  xmlChildren: XmlElement[];
  styles: StyleMap | null;
  theme: Theme | null;
  numbering: NumberingMap | null;
  rels: RelationshipMap | null;
  media: Map<string, MediaFile> | null;
  parseTable: TableParserFn;
};

const trackedChangeTypeFromXml = (localName: string): TrackedRunChange["type"] | undefined => {
  if (localName === "ins") {
    return "insertion";
  }
  if (localName === "del") {
    return "deletion";
  }
  if (localName === "moveFrom" || localName === "moveTo") {
    return localName;
  }
  return undefined;
};

const enrichTextBoxRuns = ({
  content,
  xmlChildren,
  styles,
  theme,
  numbering,
  rels,
  media,
  parseTable,
}: EnrichTextBoxRunsParams): void => {
  let parsedIndex = 0;
  let lastConsumedRun: Run | undefined;

  for (const xmlChild of xmlChildren) {
    const localName = getLocalName(xmlChild.name ?? "");
    if (localName === "pPr") {
      continue;
    }
    const trackedChangeType = trackedChangeTypeFromXml(localName);
    const parsedContent = content[parsedIndex];
    if (trackedChangeType && parsedContent?.type === trackedChangeType) {
      enrichTextBoxRuns({
        content: parsedContent.content,
        xmlChildren: getChildElements(xmlChild),
        styles,
        theme,
        numbering,
        rels,
        media,
        parseTable,
      });
    }

    if (localName === "sdt" && parsedContent?.type === "inlineSdt") {
      const properties = parsedContent.properties;
      const nestedContent: InlineSdt["content"] = [];
      let lastSegmentIndex = parsedIndex;

      for (let index = parsedIndex; index < content.length; index += 1) {
        const candidate = content[index];
        if (candidate?.type !== "inlineSdt" || candidate.properties !== properties) {
          continue;
        }
        nestedContent.push(...candidate.content);
        lastSegmentIndex = index;
      }

      const sdtContent = getChildElements(xmlChild).find(
        (child) => getLocalName(child.name ?? "") === "sdtContent",
      );
      if (sdtContent) {
        enrichTextBoxRuns({
          content: nestedContent,
          xmlChildren: getChildElements(sdtContent),
          styles,
          theme,
          numbering,
          rels,
          media,
          parseTable,
        });
      }
      parsedIndex = lastSegmentIndex + 1;
      continue;
    }

    if (localName !== "r") {
      if (parsedIndex < content.length && parsedContent?.type !== "run") {
        parsedIndex += 1;
      }
      continue;
    }

    const { textBoxDrawings, vmlTextBoxes, hasNonTextBoxContent } =
      scanRunForTextBoxDrawings(xmlChild);

    const parsedRun: Run | undefined = parsedContent?.type === "run" ? parsedContent : undefined;
    const targetRun = parsedRun ?? (hasNonTextBoxContent ? lastConsumedRun : undefined);
    const targetRunMatchesXml =
      targetRun !== undefined && (hasNonTextBoxContent || parsedRun?.content.length === 0);

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
            parseTable,
            styles,
            theme,
            numbering,
            rels,
            media,
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

      if (targetRunMatchesXml) {
        targetRun.content.push(shapeContent);
      } else {
        const newRun: Run = { type: "run", content: [shapeContent] };
        content.splice(parsedIndex, 0, newRun);
        lastConsumedRun = newRun;
        parsedIndex += 1;
      }
    }

    for (const pictEl of vmlTextBoxes) {
      const shape = parseVmlTextBoxShape(pictEl, styles, theme, numbering, rels, media, parseTable);
      if (!shape) {
        continue;
      }
      const shapeContent: ShapeContent = { type: "shape", shape };
      if (targetRunMatchesXml) {
        targetRun.content.push(shapeContent);
      } else {
        const newRun: Run = { type: "run", content: [shapeContent] };
        content.splice(parsedIndex, 0, newRun);
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
  vmlTextBoxes: XmlElement[];
  hasNonTextBoxContent: boolean;
};

const scanRunForTextBoxDrawings = (xmlRun: XmlElement): TextBoxRunScan => {
  const textBoxDrawings: XmlElement[] = [];
  const vmlTextBoxes: XmlElement[] = [];
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
    if (name === "pict") {
      if (findDeep(el, "v", "textbox")) {
        // An image-backed VML container is preserved as one raw drawing by
        // runParser. Adding an editable text-box shape here would serialize a
        // second representation beside that raw replay on every save.
        if (findDeep(el, "v", "imagedata")) {
          hasNonTextBoxContent = true;
        } else {
          vmlTextBoxes.push(el);
        }
      } else {
        hasNonTextBoxContent = true;
      }
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

  return { textBoxDrawings, vmlTextBoxes, hasNonTextBoxContent };
};

const parseVmlTextBoxShape = (
  pictEl: XmlElement,
  styles: StyleMap | null,
  theme: Theme | null,
  numbering: NumberingMap | null,
  rels: RelationshipMap | null,
  media: Map<string, MediaFile> | null,
  parseTable: TableParserFn,
): Shape | null => {
  const shapeEl = findDeep(pictEl, "v", "shape");
  const textBoxEl = shapeEl ? findDeep(shapeEl, "v", "textbox") : null;
  const contentEl = textBoxEl ? findDeep(textBoxEl, "w", "txbxContent") : null;
  if (!shapeEl || !textBoxEl || !contentEl) {
    return null;
  }

  const style = parseVmlStyle(getAttribute(shapeEl, null, "style"));
  const width = vmlLengthToPixels(style["width"]);
  const height = vmlLengthToPixels(style["height"]);
  if (width === undefined || height === undefined || width <= 0 || height <= 0) {
    return null;
  }

  const left = vmlLengthToPixels(style["margin-left"] ?? style["left"]);
  const top = vmlLengthToPixels(style["margin-top"] ?? style["top"]);
  const positioned = style["position"]?.toLowerCase() === "absolute";
  const zIndex = Number.parseInt(style["z-index"] ?? "", 10);
  const margins = parseVmlInsets(textBoxEl);
  const fill = parseVmlFill(shapeEl);
  const shape: Shape = {
    type: "shape",
    shapeType: "textBox",
    size: { width: pixelsToEmu(width), height: pixelsToEmu(height) },
    ...(fill === undefined ? {} : { fill }),
    wrap: { type: vmlWrapType(positioned, zIndex) },
    textBody: {
      content: parseTextBoxContent(
        contentEl,
        parseParagraph,
        parseTable,
        styles,
        theme,
        numbering,
        rels,
        media,
      ),
      ...(margins === undefined ? {} : { margins }),
    },
  };
  const id = getAttribute(shapeEl, null, "id");
  if (id) {
    shape.id = id;
  }
  if (positioned) {
    shape.position = {
      horizontal: {
        relativeTo: horizontalRelativeTo(style["mso-position-horizontal-relative"]),
        ...(left === undefined ? {} : { posOffset: pixelsToEmu(left) }),
      },
      vertical: {
        relativeTo: verticalRelativeTo(style["mso-position-vertical-relative"]),
        ...(top === undefined ? {} : { posOffset: pixelsToEmu(top) }),
      },
    };
  }
  return shape;
};
