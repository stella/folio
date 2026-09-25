import type {
  Hyperlink,
  Image,
  ImagePosition,
  InlineSdt,
  MediaFile,
  Paragraph,
  ParagraphContent,
  RelationshipMap,
  Run,
  Shape,
  ShapeContent,
  ShapeTextBody,
  Theme,
  TrackedRunChange,
} from "../types/document";
import { pixelsToEmu } from "../utils/units";
import { groupTextContentFingerprint, groupXmlFingerprint } from "./drawingGroupChildren";
import type { GroupTextBoxFrame } from "./drawingGroupChildren";
import { groupPreviewTextBoxes } from "./groupDrawingParser";
import type { NumberingMap } from "./numberingParser";
import { parseParagraph } from "./paragraphParser";
import type { ParseContext } from "./parseContext";
import type { PreviewLedger } from "./previewBudget";
import { consolidateParagraphContent } from "./runConsolidator";
import { captureShapeAlternateContent } from "./shapeAlternateContent";
import type { StyleMap } from "./styleParser";
import {
  getTextBoxContentElement,
  parseTextBox,
  parseTextBoxContent,
  parseTextBoxFromShape,
  scanRunForTextBoxDrawings,
} from "./textBoxParser";
import type { TableParserFn } from "./textBoxParser";
import { isVmlPictParsedByRunParser } from "./vmlImageParser";
import {
  findChildByLocalName,
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
const VML_TEXTBOX_INSET_DEFAULTS = ["0.1in", "0.05in", "0.1in", "0.05in"] as const;
const MAX_VML_TEXTBOX_INSET_TOKEN_CHARACTERS = 64;
const MAX_VML_STYLE_CHARACTERS = 16_384;

const parseVmlStyle = (value: string | null): Record<string, string> => {
  const declarations: Record<string, string> = {};
  if (!value || value.length > MAX_VML_STYLE_CHARACTERS) {
    return declarations;
  }
  for (const declaration of value.split(";")) {
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

const parseVmlInsetFields = (rawInset: string): Array<string | undefined> | undefined => {
  const fields: Array<string | undefined> = [];
  let offset = 0;
  let segmentHasField = false;

  while (offset <= rawInset.length) {
    while (/\s/u.test(rawInset.charAt(offset))) {
      offset += 1;
    }
    if (offset >= rawInset.length) {
      if (!segmentHasField) {
        fields.push(undefined);
      }
      return fields.length <= VML_TEXTBOX_INSET_DEFAULTS.length ? fields : undefined;
    }
    if (rawInset.charAt(offset) === ",") {
      if (!segmentHasField) {
        fields.push(undefined);
      }
      if (fields.length > VML_TEXTBOX_INSET_DEFAULTS.length) {
        return undefined;
      }
      segmentHasField = false;
      offset += 1;
      continue;
    }

    const tokenStart = offset;
    while (offset < rawInset.length && !/[,\s]/u.test(rawInset.charAt(offset))) {
      offset += 1;
      if (offset - tokenStart > MAX_VML_TEXTBOX_INSET_TOKEN_CHARACTERS) {
        return undefined;
      }
    }
    fields.push(rawInset.slice(tokenStart, offset));
    if (fields.length > VML_TEXTBOX_INSET_DEFAULTS.length) {
      return undefined;
    }
    segmentHasField = true;
  }

  return undefined;
};

const parseVmlInsets = (
  textBoxEl: XmlElement,
): { left: number; top: number; right: number; bottom: number } | undefined => {
  const rawInset = getAttribute(textBoxEl, null, "inset") ?? "";
  const fields = parseVmlInsetFields(rawInset);
  if (!fields) {
    return undefined;
  }
  const [left, top, right, bottom] = VML_TEXTBOX_INSET_DEFAULTS.map((fallback, index) =>
    vmlLengthToPixels(fields.at(index) ?? fallback),
  );
  if (left === undefined || top === undefined || right === undefined || bottom === undefined) {
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

const parseVmlTextAnchor = (value: string | undefined): ShapeTextBody["anchor"] | undefined => {
  switch (value?.toLowerCase()) {
    case "top":
      return "top";
    case "middle":
      return "middle";
    case "bottom":
      return "bottom";
    default:
      return undefined;
  }
};

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
  previews: PreviewLedger,
  context?: ParseContext,
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
    previews,
    context,
  });
  liftGroupTextBoxes(paragraph.content, {
    styles,
    theme,
    numbering,
    rels,
    media,
    parseTable,
    previews,
  });
  // Matching consumes source w:r elements by position. Merge only after that
  // source-dependent pass so consolidation cannot move a box to another run.
  paragraph.content = consolidateParagraphContent(paragraph.content);
};

// ============================================================================
// TEXT BOXES INSIDE DRAWINGML GROUPS
// ============================================================================

type GroupTextBoxParsers = VmlTextBoxShapeParsers;

type LiftableContent =
  | ParagraphContent
  | InlineSdt["content"][number]
  | TrackedRunChange["content"][number]
  | Hyperlink["children"][number];

/** Previews whose text boxes were already lifted, so a second pass adds none. */
const liftedGroupPreviews = new WeakSet<Image>();

/**
 * Lift the text boxes out of every group preview in the paragraph, each into
 * a text box shape right after the group's drawing in the same run.
 *
 * The group keeps drawing its other children; a lifted text box is laid out
 * and painted like any anchored text box, placed at the group's offset plus
 * the child's frame, and it records where in the group it came from so the
 * writer can put its text back there.
 */
const liftGroupTextBoxes = (
  content: readonly LiftableContent[],
  parsers: GroupTextBoxParsers,
): void => {
  for (const item of content) {
    switch (item.type) {
      case "run":
        liftGroupTextBoxesInRun(item, parsers);
        break;
      case "hyperlink":
        liftGroupTextBoxes(item.children, parsers);
        break;
      case "inlineSdt":
      case "insertion":
      case "deletion":
      case "moveFrom":
      case "moveTo":
        liftGroupTextBoxes(item.content, parsers);
        break;
      default:
        break;
    }
  }
};

const liftGroupTextBoxesInRun = (run: Run, parsers: GroupTextBoxParsers): void => {
  if (!run.content.some((item) => item.type === "drawing")) {
    return;
  }
  const lifted: Run["content"] = [];
  for (const item of run.content) {
    lifted.push(item);
    if (item.type !== "drawing" || item.rawXml === undefined) {
      continue;
    }
    const frames = groupPreviewTextBoxes(item.image);
    if (!frames || liftedGroupPreviews.has(item.image)) {
      continue;
    }
    liftedGroupPreviews.add(item.image);
    const group = groupXmlFingerprint(item.rawXml);
    for (const frame of frames) {
      const shape = groupTextBoxShape(item.image, frame, group, parsers);
      if (shape) {
        lifted.push({ type: "shape", shape });
      }
    }
  }
  run.content = lifted;
};

const groupTextBoxShape = (
  groupImage: Image,
  frame: GroupTextBoxFrame,
  group: string,
  { styles, theme, numbering, rels, media, parseTable, previews }: GroupTextBoxParsers,
): Shape | undefined => {
  const position = groupImage.position;
  const horizontalOffset = position?.horizontal.posOffset;
  const verticalOffset = position?.vertical.posOffset;
  if (!position || horizontalOffset === undefined || verticalOffset === undefined) {
    return undefined;
  }
  const size = { width: Math.round(frame.width), height: Math.round(frame.height) };
  const textBox = parseTextBoxFromShape(frame.wsp, size);
  const content = parseTextBoxContent(
    findChildByLocalName(findChildByLocalName(frame.wsp, "txbx"), "txbxContent"),
    parseParagraph,
    parseTable,
    styles,
    theme,
    numbering,
    rels,
    media,
    previews,
  );
  const transform =
    frame.rotation !== 0 || frame.flipH || frame.flipV
      ? {
          ...(frame.rotation !== 0 ? { rotation: frame.rotation } : {}),
          ...(frame.flipH ? { flipH: true } : {}),
          ...(frame.flipV ? { flipV: true } : {}),
        }
      : undefined;
  const shape: Shape = {
    type: "shape",
    shapeType: "textBox",
    size,
    ...(textBox?.name !== undefined ? { name: textBox.name } : {}),
    ...(textBox?.alt !== undefined ? { alt: textBox.alt } : {}),
    ...(textBox?.title !== undefined ? { title: textBox.title } : {}),
    position: {
      horizontal: {
        relativeTo: position.horizontal.relativeTo,
        posOffset: Math.round(horizontalOffset + frame.x),
      },
      vertical: {
        relativeTo: position.vertical.relativeTo,
        posOffset: Math.round(verticalOffset + frame.y),
      },
    },
    // The group owns any wrapping; its children only stack in its layer.
    wrap: { type: groupImage.wrap.type === "behind" ? "behind" : "inFront" },
    ...(groupImage.anchor !== undefined ? { anchor: { ...groupImage.anchor } } : {}),
    ...(textBox?.fill !== undefined ? { fill: textBox.fill } : {}),
    ...(textBox?.outline !== undefined ? { outline: textBox.outline } : {}),
    ...(transform !== undefined ? { transform } : {}),
    textBody: {
      content,
      ...(textBox?.autoFit !== undefined ? { autoFit: textBox.autoFit } : {}),
      ...(textBox?.textWrap !== undefined ? { textWrap: textBox.textWrap } : {}),
      ...(textBox?.verticalAlign !== undefined ? { anchor: textBox.verticalAlign } : {}),
      ...(textBox?.margins !== undefined ? { margins: textBox.margins } : {}),
    },
    groupChild: {
      path: frame.path,
      group,
      content: groupTextContentFingerprint(content),
    },
  };
  if (textBox?.id) {
    shape.id = textBox.id;
  }
  return shape;
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
  previews: PreviewLedger;
  /** Absent when the caller has no warning collector; see `blockContentParser`. */
  context: ParseContext | undefined;
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
  previews,
  context,
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
        previews,
        context,
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
          previews,
          context,
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

    const { textBoxDrawings, vmlTextBoxes, hasNonTextBoxContent } = scanRunForTextBoxDrawings({
      xmlRun: xmlChild,
      claimedByRunParser: (pictElement) => isVmlPictParsedByRunParser(pictElement, rels, media),
    });

    const parsedRun: Run | undefined = parsedContent?.type === "run" ? parsedContent : undefined;
    const targetRun = parsedRun ?? (hasNonTextBoxContent ? lastConsumedRun : undefined);
    const targetRunMatchesXml =
      targetRun !== undefined && (hasNonTextBoxContent || parsedRun?.content.length === 0);
    // This `w:r` carried nothing but the text box, and the run that waited for
    // it is the one at `parsedIndex`. Filling it consumes that position, so the
    // next `w:r` must look past it; leaving the index where it was made a
    // second box insert itself in front of the first.
    const fillsEmptyCarrier =
      targetRunMatchesXml && !hasNonTextBoxContent && parsedRun !== undefined;

    for (const { drawing: runEl, alternateContent: source } of textBoxDrawings) {
      const textBox = parseTextBox(runEl, context);
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
            previews,
          );
        }
      }

      const shape: Shape = {
        type: "shape",
        shapeType: "textBox",
        size: textBox.size,
        ...(textBox.name !== undefined ? { name: textBox.name } : {}),
        ...(textBox.alt !== undefined ? { alt: textBox.alt } : {}),
        ...(textBox.title !== undefined ? { title: textBox.title } : {}),
        ...(textBox.position !== undefined ? { position: textBox.position } : {}),
        ...(textBox.wrap !== undefined ? { wrap: textBox.wrap } : {}),
        ...(textBox.anchor !== undefined ? { anchor: textBox.anchor } : {}),
        ...(textBox.fill !== undefined ? { fill: textBox.fill } : {}),
        ...(textBox.outline !== undefined ? { outline: textBox.outline } : {}),
        textBody: {
          content: textBox.content,
          ...(textBox.autoFit !== undefined ? { autoFit: textBox.autoFit } : {}),
          ...(textBox.textWrap !== undefined ? { textWrap: textBox.textWrap } : {}),
          ...(textBox.verticalAlign !== undefined ? { anchor: textBox.verticalAlign } : {}),
          ...(textBox.margins !== undefined ? { margins: textBox.margins } : {}),
          ...(textBox.wordArt !== undefined ? { wordArt: textBox.wordArt } : {}),
        },
      };
      if (textBox.id) {
        shape.id = textBox.id;
      }

      const alternateContent =
        source &&
        captureShapeAlternateContent({
          shape,
          alternateContent: source.element,
          branch: source.branch,
          drawing: runEl,
        });
      const shapeContent: ShapeContent = {
        type: "shape",
        shape,
        ...(alternateContent ? { alternateContent } : {}),
      };

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
      const shape = parseVmlTextBoxShape(pictEl, {
        styles,
        theme,
        numbering,
        rels,
        media,
        parseTable,
        previews,
      });
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

    if ((hasNonTextBoxContent || fillsEmptyCarrier) && parsedRun) {
      lastConsumedRun = parsedRun;
      parsedIndex += 1;
    }
  }
};

type VmlTextBoxShapeParsers = {
  styles: StyleMap | null;
  theme: Theme | null;
  numbering: NumberingMap | null;
  rels: RelationshipMap | null;
  media: Map<string, MediaFile> | null;
  parseTable: TableParserFn;
  previews: PreviewLedger;
};

const parseVmlTextBoxShape = (
  pictEl: XmlElement,
  { styles, theme, numbering, rels, media, parseTable, previews }: VmlTextBoxShapeParsers,
): Shape | null => {
  const shapeEl = findDeep(pictEl, "v", "shape");
  const textBoxEl = shapeEl ? findDeep(shapeEl, "v", "textbox") : null;
  const contentEl = textBoxEl ? findDeep(textBoxEl, "w", "txbxContent") : null;
  if (!shapeEl || !textBoxEl || !contentEl) {
    return null;
  }

  const style = parseVmlStyle(getAttribute(shapeEl, null, "style"));
  const textBoxStyle = parseVmlStyle(getAttribute(textBoxEl, null, "style"));
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
  const anchor = parseVmlTextAnchor(textBoxStyle["v-text-anchor"]);
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
        previews,
      ),
      ...(margins === undefined ? {} : { margins }),
      ...(anchor === undefined ? {} : { anchor }),
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
        posOffset: pixelsToEmu(left ?? 0),
      },
      vertical: {
        relativeTo: verticalRelativeTo(style["mso-position-vertical-relative"]),
        posOffset: pixelsToEmu(top ?? 0),
      },
    };
  }
  return shape;
};
