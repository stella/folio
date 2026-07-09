import type { Image } from "../types/document";
import { emuToPixels } from "../utils/units";
import { parseImage } from "./imageParser";
import {
  findAllDeep,
  findChildByLocalName,
  findChildrenByLocalName,
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
const MAX_PATH_COMMANDS = 10_000;
const MAX_TEXT_CHARACTERS = 20_000;
const MAX_SVG_CHARACTERS = 1_000_000;

const escapeXml = (value: string): string =>
  value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&apos;");

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

const childTransform = (wsp: XmlElement): { x: number; y: number; width: number; height: number } => {
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
  const solidFill = findChildByLocalName(parent, "solidFill");
  const srgb = findChildByLocalName(solidFill, "srgbClr");
  const value = getAttribute(srgb, null, "val");
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

const renderGeometry = (wsp: XmlElement): string => {
  const { x, y, width, height } = childTransform(wsp);
  if (width <= 0 || height <= 0) {
    return "";
  }
  const spPr = findChildByLocalName(wsp, "spPr");
  const fill = colorFrom(spPr, "none");
  const line = findChildByLocalName(spPr, "ln");
  const stroke = colorFrom(line, "none");
  const strokeWidth = line
    ? (parseNumericAttribute(line, null, "w") ?? DEFAULT_LINE_WIDTH_EMU)
    : 0;
  const paths = findAllDeep(findChildByLocalName(spPr, "custGeom"), "a", "path");
  if (paths.length === 0) {
    return `<rect x="${x}" y="${y}" width="${width}" height="${height}" fill="${fill === "none" ? "none" : `#${fill}`}" stroke="${stroke === "none" ? "none" : `#${stroke}`}" stroke-width="${strokeWidth}"/>`;
  }
  return paths
    .map((path) => {
      const viewWidth = numericAttr(path, "w") || width;
      const viewHeight = numericAttr(path, "h") || height;
      const d = pathData(path);
      if (!d) {
        return "";
      }
      return `<path d="${d}" transform="translate(${x} ${y}) scale(${width / viewWidth} ${height / viewHeight})" fill="${fill === "none" ? "none" : `#${fill}`}" stroke="${stroke === "none" ? "none" : `#${stroke}`}" stroke-width="${strokeWidth}"/>`;
    })
    .join("");
};

const renderTextBox = (wsp: XmlElement): string => {
  const { x, y, width, height } = childTransform(wsp);
  if (width <= 0 || height <= 0) {
    return "";
  }
  const paragraphs = findAllDeep(wsp, "w", "p");
  const firstSize = findAllDeep(wsp, "w", "sz").at(0);
  const halfPoints = numericAttr(firstSize ?? null, "val") || DEFAULT_FONT_HALF_POINTS;
  const color = colorFrom(findAllDeep(wsp, "w", "color").at(0) ?? null, DEFAULT_TEXT_COLOR);
  const fontSize = halfPoints * HALF_POINT_TO_EMU;
  const maxCharacters = Math.max(1, Math.floor(width / (fontSize * 0.38)));
  const lines = paragraphs
    .flatMap((paragraph) =>
      wrapLine(getTextContent(paragraph).slice(0, MAX_TEXT_CHARACTERS), maxCharacters),
    )
    .map(escapeXml);
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

const createSvg = (group: XmlElement, width: number, height: number): string => {
  const children = findChildrenByLocalName(group, "wsp").slice(0, MAX_GROUP_SHAPES);
  const content = children
    .map((wsp) =>
      findChildByLocalName(wsp, "txbx") ? renderTextBox(wsp) : renderGeometry(wsp),
    )
    .join("");
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${width} ${height}" width="${emuToPixels(width)}" height="${emuToPixels(height)}">${content}</svg>`;
};

/** Parse a WordprocessingGroup drawing into a safe SVG-backed image preview. */
export const parseGroupDrawing = (drawing: XmlElement): Image | null => {
  const graphicData = findAllDeep(drawing, "a", "graphicData").at(0);
  const group = findChildByLocalName(graphicData ?? null, "wgp");
  if (!group) {
    return null;
  }
  const image = parseImage(drawing, undefined, undefined);
  if (!image || image.size.width <= 0 || image.size.height <= 0) {
    return null;
  }
  const svg = createSvg(group, image.size.width, image.size.height);
  if (svg.length > MAX_SVG_CHARACTERS) {
    return null;
  }
  image.src = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`;
  image.mimeType = "image/svg+xml";
  image.filename = "wordprocessing-group.svg";
  return image;
};
