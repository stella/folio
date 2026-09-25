import { escapeXmlAttribute, escapeXmlText } from "@stll/docx-core";

import type { Image, MediaFile, RelationshipMap } from "../types/document";
import { emuToPixels } from "../utils/units";
import { collectGroupTextBoxes, readTransform } from "./drawingGroupChildren";
import type { GroupTextBoxFrame } from "./drawingGroupChildren";
import { parseImage, resolveImageData } from "./imageParser";
import type { PreviewLedger } from "./previewBudget";
import {
  findAllDeep,
  findChildByLocalName,
  findChildrenByLocalName,
  getChildElements,
  getAttribute,
  getLocalName,
  getTextContent,
  parseNumericAttribute,
} from "./xmlParser";
import type { XmlElement } from "./xmlParser";

const HEX_COLOR = /^[0-9A-Fa-f]{6}$/u;
const DEFAULT_TEXT_COLOR = "000000";
const DEFAULT_FONT_HALF_POINTS = 22;
const DEFAULT_LINE_WIDTH_EMU = 9_525;
const HALF_POINT_TO_EMU = 6_350;
const MAX_GROUP_SHAPES = 256;
const MAX_GROUP_DEPTH = 16;
const CROP_SCALE = 100_000;
const MAX_PATH_COMMANDS = 10_000;
const MAX_TEXT_CHARACTERS = 20_000;
const MAX_SVG_CHARACTERS = 1_000_000;

const numericAttr = (element: XmlElement | null, name: string): number => {
  const direct = parseNumericAttribute(element, null, name);
  if (direct !== undefined) {
    return direct;
  }
  const wordValue = getAttribute(element, "w", name);
  if (!wordValue) {
    return 0;
  }
  const parsed = Number.parseInt(wordValue, 10);
  return Number.isNaN(parsed) ? 0 : parsed;
};

const childTransform = (
  wsp: XmlElement,
): { x: number; y: number; width: number; height: number } => {
  const spPr = findChildByLocalName(wsp, "spPr");
  const xfrm = findChildByLocalName(spPr, "xfrm");
  const off = findChildByLocalName(xfrm, "off");
  const ext = findChildByLocalName(xfrm, "ext");
  return {
    x: numericAttr(off, "x"),
    y: numericAttr(off, "y"),
    width: numericAttr(ext, "cx"),
    height: numericAttr(ext, "cy"),
  };
};

const colorFrom = (parent: XmlElement | null, fallback?: string): string | undefined => {
  if (parent && getLocalName(parent.name ?? "") === "color") {
    const value = getAttribute(parent, "w", "val");
    if (value && HEX_COLOR.test(value)) {
      return value.toUpperCase();
    }
  }
  if (findChildByLocalName(parent, "noFill")) {
    return "none";
  }
  const solidFill = findChildByLocalName(parent, "solidFill");
  const srgb = findChildByLocalName(solidFill, "srgbClr");
  const referencedSrgb = findChildByLocalName(parent, "srgbClr");
  const value = getAttribute(srgb ?? referencedSrgb, null, "val");
  if (value && HEX_COLOR.test(value)) {
    return value.toUpperCase();
  }
  return fallback;
};

const wrapLine = (line: string, maxCharacters: number): string[] => {
  const words = line.trim().split(/\s+/u);
  const lines: string[] = [];
  let current = "";
  for (const word of words) {
    const candidate = current ? `${current} ${word}` : word;
    if (current && candidate.length > maxCharacters) {
      lines.push(current);
      current = word;
    } else {
      current = candidate;
    }
  }
  if (current) {
    lines.push(current);
  }
  return lines;
};

const pathData = (path: XmlElement): string => {
  const commands: string[] = [];
  for (const command of path.elements ?? []) {
    if (commands.length >= MAX_PATH_COMMANDS) {
      break;
    }
    if (command.type !== "element") {
      continue;
    }
    const point = findChildByLocalName(command, "pt");
    const x = numericAttr(point, "x");
    const y = numericAttr(point, "y");
    const name = command.name?.split(":").at(-1);
    if (name === "moveTo") {
      commands.push(`M ${x} ${y}`);
    } else if (name === "lnTo") {
      commands.push(`L ${x} ${y}`);
    } else if (name === "quadBezTo") {
      const points = findChildrenByLocalName(command, "pt");
      if (points.length >= 2) {
        commands.push(
          `Q ${numericAttr(points[0] ?? null, "x")} ${numericAttr(points[0] ?? null, "y")} ${numericAttr(points[1] ?? null, "x")} ${numericAttr(points[1] ?? null, "y")}`,
        );
      }
    } else if (name === "cubicBezTo") {
      const points = findChildrenByLocalName(command, "pt");
      if (points.length >= 3) {
        commands.push(
          `C ${numericAttr(points[0] ?? null, "x")} ${numericAttr(points[0] ?? null, "y")} ${numericAttr(points[1] ?? null, "x")} ${numericAttr(points[1] ?? null, "y")} ${numericAttr(points[2] ?? null, "x")} ${numericAttr(points[2] ?? null, "y")}`,
        );
      }
    } else if (name === "close") {
      commands.push("Z");
    }
  }
  return commands.join(" ");
};

/**
 * A child's own `a:xfrm` rotation and flips, applied about its centre in the
 * coordinate space it is drawn in (ECMA-376 §20.1.7.6).
 */
const withChildTransform = (element: XmlElement, piece: string): string => {
  if (!piece) {
    return piece;
  }
  const spPr = findChildByLocalName(element, "spPr");
  const transform = readTransform(findChildByLocalName(spPr, "xfrm"));
  return wrapTransformed(transform, piece);
};

const wrapTransformed = (
  transform: { x: number; y: number; width: number; height: number } & Pick<
    ReturnType<typeof readTransform>,
    "rotation" | "flipH" | "flipV"
  >,
  piece: string,
): string => {
  if (transform.rotation === 0 && !transform.flipH && !transform.flipV) {
    return piece;
  }
  const centerX = transform.x + transform.width / 2;
  const centerY = transform.y + transform.height / 2;
  const operations: string[] = [];
  if (transform.rotation !== 0) {
    operations.push(`rotate(${transform.rotation} ${centerX} ${centerY})`);
  }
  if (transform.flipH || transform.flipV) {
    operations.push(
      `translate(${centerX} ${centerY}) scale(${transform.flipH ? -1 : 1} ${transform.flipV ? -1 : 1}) translate(${-centerX} ${-centerY})`,
    );
  }
  return `<g transform="${operations.join(" ")}">${piece}</g>`;
};

/**
 * `unitScale` is how many EMUs one unit of the child coordinate space spans.
 * Line widths are authored in EMUs, so they are divided by it to be drawn in
 * the space the geometry is in.
 */
const renderGeometry = (wsp: XmlElement, unitScale: number): string => {
  const { x, y, width, height } = childTransform(wsp);
  const spPr = findChildByLocalName(wsp, "spPr");
  const style = findChildByLocalName(wsp, "style");
  const styleStroke = colorFrom(findChildByLocalName(style, "lnRef"), "none");
  const fill = colorFrom(spPr, "none");
  const line = findChildByLocalName(spPr, "ln");
  const stroke = colorFrom(line, styleStroke);
  const strokeWidth =
    (line ? (parseNumericAttribute(line, null, "w") ?? DEFAULT_LINE_WIDTH_EMU) : 0) / unitScale;
  const strokeColor = stroke === "none" ? "none" : `#${stroke}`;
  if (width <= 0 || height <= 0) {
    // A straight line or connector is a frame with no width or no height; the
    // line runs corner to corner of it, so a zero extent still draws.
    if (width < 0 || height < 0 || (width === 0 && height === 0) || strokeColor === "none") {
      return "";
    }
    return `<line x1="${x}" y1="${y}" x2="${x + width}" y2="${y + height}" stroke="${strokeColor}" stroke-width="${strokeWidth}"/>`;
  }
  const paths = findAllDeep(findChildByLocalName(spPr, "custGeom"), "a", "path");
  if (paths.length === 0) {
    return `<rect x="${x}" y="${y}" width="${width}" height="${height}" fill="${fill === "none" ? "none" : `#${fill}`}" stroke="${strokeColor}" stroke-width="${strokeWidth}"/>`;
  }
  return paths
    .map((path) => {
      const viewWidth = numericAttr(path, "w") || width;
      const viewHeight = numericAttr(path, "h") || height;
      const d = pathData(path);
      if (!d) {
        return "";
      }
      return `<path d="${d}" transform="translate(${x} ${y}) scale(${width / viewWidth} ${height / viewHeight})" fill="${fill === "none" ? "none" : `#${fill}`}" stroke="${strokeColor}" stroke-width="${strokeWidth}"/>`;
    })
    .join("");
};

/**
 * An approximate rendering of a text box the layout does not host (an inline
 * group, or one placed by alignment rather than offset). Font sizes are in
 * EMUs and the drawing is in child units, hence `unitScale`.
 */
const renderTextBox = (wsp: XmlElement, unitScale: number): string => {
  const { x, y, width, height } = childTransform(wsp);
  if (width <= 0 || height <= 0) {
    return "";
  }
  const paragraphs = findAllDeep(wsp, "w", "p");
  const firstSize = findAllDeep(wsp, "w", "sz").at(0);
  const halfPoints = numericAttr(firstSize ?? null, "val") || DEFAULT_FONT_HALF_POINTS;
  const color = colorFrom(findAllDeep(wsp, "w", "color").at(0) ?? null, DEFAULT_TEXT_COLOR);
  const fontSize = (halfPoints * HALF_POINT_TO_EMU) / unitScale;
  const maxCharacters = Math.max(1, Math.floor(width / (fontSize * 0.38)));
  const lines = paragraphs
    .flatMap((paragraph) =>
      wrapLine(getTextContent(paragraph).slice(0, MAX_TEXT_CHARACTERS), maxCharacters),
    )
    .map(escapeXmlText);
  if (lines.length === 0) {
    return "";
  }
  const lineHeight = fontSize * 1.15;
  const svgFontSize = 1_000;
  const scale = fontSize / svgFontSize;
  const lineStep = (lineHeight / fontSize) * svgFontSize;
  const tspans = lines
    .map((line, index) => `<tspan x="0" dy="${index === 0 ? 0 : lineStep}">${line}</tspan>`)
    .join("");
  return `<text x="0" y="${svgFontSize}" transform="translate(${x} ${y}) scale(${scale})" font-family="Arial, sans-serif" font-size="${svgFontSize}" fill="#${color}">${tspans}</text>`;
};

const renderPicture = (
  picture: XmlElement,
  index: number,
  rels: RelationshipMap | undefined,
  media: Map<string, MediaFile> | undefined,
): string => {
  const { x, y, width, height } = childTransform(picture);
  if (width <= 0 || height <= 0) {
    return "";
  }
  const blipFill = findChildByLocalName(picture, "blipFill");
  const blip = findChildByLocalName(blipFill, "blip");
  const rId = getAttribute(blip, "r", "embed") ?? getAttribute(blip, "r", "link") ?? undefined;
  const { src } = resolveImageData(rId, rels, media);
  if (!src) {
    return "";
  }

  const sourceRect = findChildByLocalName(blipFill, "srcRect");
  const left = Math.max(0, numericAttr(sourceRect, "l")) / CROP_SCALE;
  const top = Math.max(0, numericAttr(sourceRect, "t")) / CROP_SCALE;
  const right = Math.max(0, numericAttr(sourceRect, "r")) / CROP_SCALE;
  const bottom = Math.max(0, numericAttr(sourceRect, "b")) / CROP_SCALE;
  const visibleWidth = 1 - left - right;
  const visibleHeight = 1 - top - bottom;
  if (visibleWidth <= 0 || visibleHeight <= 0) {
    return "";
  }

  const imageX = x - (width * left) / visibleWidth;
  const imageY = y - (height * top) / visibleHeight;
  const imageWidth = width / visibleWidth;
  const imageHeight = height / visibleHeight;
  const image = `<image x="${imageX}" y="${imageY}" width="${imageWidth}" height="${imageHeight}" href="${escapeXmlAttribute(src)}" preserveAspectRatio="none"/>`;
  if (left === 0 && top === 0 && right === 0 && bottom === 0) {
    return image;
  }
  const clipId = `group-picture-${index}`;
  return `<defs><clipPath id="${clipId}"><rect x="${x}" y="${y}" width="${width}" height="${height}"/></clipPath></defs><g clip-path="url(#${clipId})">${image}</g>`;
};

const groupViewBox = (
  group: XmlElement,
  width: number,
  height: number,
): { x: number; y: number; width: number; height: number } => {
  const transform = readTransform(
    findChildByLocalName(findChildByLocalName(group, "grpSpPr"), "xfrm"),
  );
  return {
    x: transform.childX,
    y: transform.childY,
    width: transform.childWidth > 0 ? transform.childWidth : width,
    height: transform.childHeight > 0 ? transform.childHeight : height,
  };
};

/**
 * A nested group's child space drawn inside its frame in the parent's space,
 * then turned by the group's own rotation and flips.
 */
const nestedGroupTransform = (transform: ReturnType<typeof readTransform>): string => {
  const scaleX = transform.childWidth > 0 ? transform.width / transform.childWidth : 1;
  const scaleY = transform.childHeight > 0 ? transform.height / transform.childHeight : 1;
  return `translate(${transform.x} ${transform.y}) scale(${scaleX} ${scaleY}) translate(${-transform.childX} ${-transform.childY})`;
};

/** EMUs per child unit, per axis. */
type UnitScale = { x: number; y: number };

const averageScale = (scale: UnitScale): number => (scale.x + scale.y) / 2 || 1;

const createSvg = (
  group: XmlElement,
  width: number,
  height: number,
  rels: RelationshipMap | undefined,
  media: Map<string, MediaFile> | undefined,
  hostedTextBoxes: ReadonlySet<XmlElement>,
): string | null => {
  // Count every character as it is produced and bail as soon as the budget
  // is exceeded, rather than joining every child's markup (each of which can
  // embed a large media data URL) and checking the total afterward — that
  // would let a hostile group of otherwise-individually-bounded shapes build
  // a multi-GB transient string before ever being rejected.
  let total = 0;
  let shapes = 0;
  let serial = 0;
  /** One container's markup, or null once the budget is spent. */
  const render = (container: XmlElement, scale: UnitScale, depth: number): string | null => {
    let content = "";
    for (const child of getChildElements(container)) {
      const index = serial;
      serial += 1;
      if (shapes >= MAX_GROUP_SHAPES) {
        break;
      }
      const localName = getLocalName(child.name ?? "");
      let piece: string;
      if (localName === "pic") {
        shapes += 1;
        piece = withChildTransform(child, renderPicture(child, index, rels, media));
      } else if (localName === "wsp") {
        shapes += 1;
        if (hostedTextBoxes.has(child)) {
          // The layout draws this one, frame and all, as a text box.
          continue;
        }
        const unitScale = averageScale(scale);
        piece = withChildTransform(
          child,
          findChildByLocalName(child, "txbx")
            ? `${renderGeometry(child, unitScale)}${renderTextBox(child, unitScale)}`
            : renderGeometry(child, unitScale),
        );
      } else if (localName === "grpSp" && depth < MAX_GROUP_DEPTH) {
        const transform = readTransform(
          findChildByLocalName(findChildByLocalName(child, "grpSpPr"), "xfrm"),
        );
        const scaleX = transform.childWidth > 0 ? transform.width / transform.childWidth : 1;
        const scaleY = transform.childHeight > 0 ? transform.height / transform.childHeight : 1;
        const nested = render(child, { x: scale.x * scaleX, y: scale.y * scaleY }, depth + 1);
        if (nested === null) {
          return null;
        }
        piece = nested
          ? wrapTransformed(
              transform,
              `<g transform="${nestedGroupTransform(transform)}">${nested}</g>`,
            )
          : "";
        // The nested markup was counted as it was produced.
        total -= nested.length;
      } else {
        continue;
      }
      content += piece;
      total += piece.length;
      if (total > MAX_SVG_CHARACTERS) {
        return null;
      }
    }
    return content;
  };
  const viewBox = groupViewBox(group, width, height);
  const content = render(group, { x: width / viewBox.width, y: height / viewBox.height }, 0);
  if (content === null || (!content && hostedTextBoxes.size === 0)) {
    return null;
  }
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="${viewBox.x} ${viewBox.y} ${viewBox.width} ${viewBox.height}" width="${emuToPixels(width)}" height="${emuToPixels(height)}">${content}</svg>`;
};

const groupElement = (drawing: XmlElement): XmlElement | null => {
  const graphicData = findAllDeep(drawing, "a", "graphicData").at(0);
  return findChildByLocalName(graphicData ?? null, "wgp");
};

/**
 * Whether the layout can place the group's text boxes itself. A floating group
 * positioned by offsets on both axes puts each child at the group's offset
 * plus the child's frame, which a text box anchored beside the group can state
 * exactly; an inline group or one placed by alignment has no such offset.
 */
const hostsTextBoxes = (image: Image): boolean => {
  const position = image.position;
  return (
    image.wrap.type !== "inline" &&
    image.anchor?.useSimplePosition !== true &&
    position !== undefined &&
    position.horizontal.alignment === undefined &&
    position.horizontal.posOffset !== undefined &&
    position.vertical.alignment === undefined &&
    position.vertical.posOffset !== undefined
  );
};

/** The text boxes each group preview leaves to the layout, keyed by the preview. */
const hostedTextBoxFrames = new WeakMap<Image, readonly GroupTextBoxFrame[]>();

/**
 * The text boxes a group preview does not draw, which the paragraph parser
 * lifts out as text boxes of their own. Undefined for any other image.
 */
export const groupPreviewTextBoxes = (image: Image): readonly GroupTextBoxFrame[] | undefined =>
  hostedTextBoxFrames.get(image);

/**
 * Whether a `w:drawing` carries a WordprocessingGroup payload. A group this
 * module declines to rasterize has no editable projection either — the shape
 * model holds one shape, not a group — so the caller must preserve it raw.
 */
export const isGroupDrawing = (drawing: XmlElement): boolean => groupElement(drawing) !== null;

/** Parse a WordprocessingGroup drawing into a safe SVG-backed image preview. */
export const parseGroupDrawing = (
  drawing: XmlElement,
  previews: PreviewLedger,
  rels?: RelationshipMap,
  media?: Map<string, MediaFile>,
): Image | null => {
  const group = groupElement(drawing);
  if (!group) {
    return null;
  }
  const image = parseImage(drawing, undefined, undefined);
  if (!image || image.size.width <= 0 || image.size.height <= 0) {
    return null;
  }
  const hosted = hostsTextBoxes(image)
    ? collectGroupTextBoxes(group, image.size.width, image.size.height)
    : [];
  const svg = createSvg(
    group,
    image.size.width,
    image.size.height,
    rels,
    media,
    new Set(hosted.map(({ wsp }) => wsp)),
  );
  if (!svg || svg.length > MAX_SVG_CHARACTERS) {
    return null;
  }
  // The anchor is read without relationships, so it names no picture: it is
  // the frame the preview is drawn over.
  const { rId: _rId, src: _src, mimeType: _mimeType, filename: _filename, ...frame } = image;
  const preview = previews.svgImage("wpGroup", svg, frame);
  if (hosted.length > 0) {
    hostedTextBoxFrames.set(preview, hosted);
  }
  return preview;
};
