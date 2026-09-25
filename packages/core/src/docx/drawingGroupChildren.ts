/**
 * The text boxes of a DrawingML group (`wpg:wgp`), read as frames folio can lay
 * out, and written back into the group's authored XML.
 *
 * A group states a child coordinate space (`a:chOff` / `a:chExt`) that its
 * own frame (`a:off` / `a:ext`) maps onto, and every child's `a:xfrm` is in
 * that space; a nested `wpg:grpSp` repeats the mapping one level down
 * (ECMA-376 §20.1.7.5). Composing the mappings gives each `wps:wsp` a frame
 * in EMUs relative to the group's own top-left corner. Only the frame is
 * mapped: the text inside keeps its authored size and the body insets stay in
 * EMUs.
 */

import { canonicalJson } from "../utils/canonicalJson";
import { captureVerbatimXml } from "./verbatimCapture";
import {
  findDeep,
  findChildByLocalName,
  getAttribute,
  getChildElements,
  getLocalName,
  parseNumericAttribute,
  parseXml,
} from "./xmlParser";
import type { XmlElement } from "./xmlParser";

/** `a:xfrm@rot` is in 60,000ths of a degree. */
const ROTATION_UNITS_PER_DEGREE = 60_000;
/** Deeper nesting than this is not a drawing anyone authored. */
const MAX_GROUP_DEPTH = 16;
const MAX_GROUP_TEXT_BOXES = 256;

/** Child coordinates to group EMUs, per axis: `emu = child * scale + offset`. */
type AxisMap = { scaleX: number; scaleY: number; offsetX: number; offsetY: number };

export type GroupTextBoxFrame = {
  /** Element indices from the `wpg:wgp` down to the `wps:wsp`. */
  path: number[];
  wsp: XmlElement;
  /** EMUs from the group's top-left corner. */
  x: number;
  y: number;
  width: number;
  height: number;
  /** Degrees clockwise. */
  rotation: number;
  flipH: boolean;
  flipV: boolean;
};

const numberAttribute = (element: XmlElement | null, name: string): number =>
  parseNumericAttribute(element, null, name) ?? 0;

const onOffAttribute = (element: XmlElement | null, name: string): boolean => {
  const value = getAttribute(element, null, name);
  return value === "1" || value === "true" || value === "on";
};

type Transform2D = {
  x: number;
  y: number;
  width: number;
  height: number;
  childX: number;
  childY: number;
  childWidth: number;
  childHeight: number;
  rotation: number;
  flipH: boolean;
  flipV: boolean;
};

/** A `CT_GroupTransform2D` or `CT_Transform2D`; absent children read as zero. */
export const readTransform = (xfrm: XmlElement | null): Transform2D => {
  const offset = findChildByLocalName(xfrm, "off");
  const extent = findChildByLocalName(xfrm, "ext");
  const childOffset = findChildByLocalName(xfrm, "chOff");
  const childExtent = findChildByLocalName(xfrm, "chExt");
  return {
    x: numberAttribute(offset, "x"),
    y: numberAttribute(offset, "y"),
    width: numberAttribute(extent, "cx"),
    height: numberAttribute(extent, "cy"),
    childX: numberAttribute(childOffset, "x"),
    childY: numberAttribute(childOffset, "y"),
    childWidth: numberAttribute(childExtent, "cx"),
    childHeight: numberAttribute(childExtent, "cy"),
    rotation: numberAttribute(xfrm, "rot") / ROTATION_UNITS_PER_DEGREE,
    flipH: onOffAttribute(xfrm, "flipH"),
    flipV: onOffAttribute(xfrm, "flipV"),
  };
};

const groupTransform = (group: XmlElement): Transform2D =>
  readTransform(findChildByLocalName(findChildByLocalName(group, "grpSpPr"), "xfrm"));

const shapeTransform = (shape: XmlElement): Transform2D =>
  readTransform(findChildByLocalName(findChildByLocalName(shape, "spPr"), "xfrm"));

/**
 * The top-level mapping: the group's child space stretched over the drawing's
 * extent. A missing or empty child extent means the child space is the
 * drawing's own.
 */
const rootMap = (group: XmlElement, width: number, height: number): AxisMap => {
  const transform = groupTransform(group);
  const scaleX = transform.childWidth > 0 ? width / transform.childWidth : 1;
  const scaleY = transform.childHeight > 0 ? height / transform.childHeight : 1;
  return {
    scaleX,
    scaleY,
    offsetX: -transform.childX * scaleX,
    offsetY: -transform.childY * scaleY,
  };
};

/** A nested group's child space, mapped through its frame in the parent's space. */
const nestedMap = (parent: AxisMap, transform: Transform2D): AxisMap => {
  const scaleX = transform.childWidth > 0 ? transform.width / transform.childWidth : 1;
  const scaleY = transform.childHeight > 0 ? transform.height / transform.childHeight : 1;
  return {
    scaleX: parent.scaleX * scaleX,
    scaleY: parent.scaleY * scaleY,
    offsetX: parent.scaleX * (transform.x - transform.childX * scaleX) + parent.offsetX,
    offsetY: parent.scaleY * (transform.y - transform.childY * scaleY) + parent.offsetY,
  };
};

const hasTextBox = (wsp: XmlElement): boolean =>
  findChildByLocalName(findChildByLocalName(wsp, "txbx"), "txbxContent") !== null;

/**
 * Every `wps:wsp` with a text box in the group, nested groups included, with
 * its frame in EMUs from the group's top-left corner.
 *
 * A nested group that is rotated or flipped turns its children's frames into
 * something other than a rectangle in the drawing's axes, so its text boxes
 * are left out and stay with the preview.
 */
export const collectGroupTextBoxes = (
  group: XmlElement,
  width: number,
  height: number,
): GroupTextBoxFrame[] => {
  const frames: GroupTextBoxFrame[] = [];
  const visit = (container: XmlElement, map: AxisMap, path: number[], depth: number): void => {
    for (const [index, child] of getChildElements(container).entries()) {
      if (frames.length >= MAX_GROUP_TEXT_BOXES) {
        return;
      }
      const localName = getLocalName(child.name ?? "");
      if (localName === "grpSp" && depth < MAX_GROUP_DEPTH) {
        const transform = groupTransform(child);
        if (transform.rotation === 0 && !transform.flipH && !transform.flipV) {
          visit(child, nestedMap(map, transform), [...path, index], depth + 1);
        }
        continue;
      }
      if (localName !== "wsp" || !hasTextBox(child)) {
        continue;
      }
      const transform = shapeTransform(child);
      const frameWidth = transform.width * map.scaleX;
      const frameHeight = transform.height * map.scaleY;
      if (!(frameWidth > 0 && frameHeight > 0)) {
        continue;
      }
      frames.push({
        path: [...path, index],
        wsp: child,
        x: transform.x * map.scaleX + map.offsetX,
        y: transform.y * map.scaleY + map.offsetY,
        width: frameWidth,
        height: frameHeight,
        rotation: transform.rotation,
        flipH: transform.flipH,
        flipV: transform.flipV,
      });
    }
  };
  visit(group, rootMap(group, width, height), [], 0);
  return frames;
};

// ============================================================================
// FINGERPRINTS
// ============================================================================

/** Two FNV-1a passes with different offsets: a short, stable identity. */
const fingerprint = (text: string): string => {
  let first = 0x81_1c_9d_c5;
  let second = 0x01_00_01_93;
  for (let index = 0; index < text.length; index += 1) {
    const code = text.charCodeAt(index);
    first = Math.imul(first ^ code, 0x01_00_01_93);
    second = Math.imul(second ^ code, 0x05_bd_1e_95);
  }
  return `${(first >>> 0).toString(16).padStart(8, "0")}${(second >>> 0).toString(16).padStart(8, "0")}`;
};

/** Identity of a group drawing's authored XML. */
export const groupXmlFingerprint = (rawXml: string): string => fingerprint(rawXml);

/** Identity of a text box's content, compared on save to tell an edit. */
export const groupTextContentFingerprint = (content: unknown): string =>
  fingerprint(canonicalJson(content));

// ============================================================================
// WRITING BACK
// ============================================================================

export type GroupTextBoxEdit = {
  path: readonly number[];
  /** Serialized `w:txbxContent` children. */
  contentXml: string;
};

/** The group a captured drawing carries; a Fallback branch holds none. */
const findGroup = (root: XmlElement): XmlElement | undefined =>
  findDeep(root, "wpg", "wgp") ?? undefined;

const childAtPath = (group: XmlElement, path: readonly number[]): XmlElement | undefined => {
  let current: XmlElement | undefined = group;
  for (const index of path) {
    current = current ? getChildElements(current)[index] : undefined;
  }
  return current;
};

/**
 * The text box content element a path names, or undefined when the path does
 * not lead to a `wps:wsp` holding one.
 */
const textBoxContentAt = (group: XmlElement, path: readonly number[]): XmlElement | undefined => {
  const wsp = childAtPath(group, path);
  if (!wsp || getLocalName(wsp.name ?? "") !== "wsp") {
    return undefined;
  }
  return findChildByLocalName(findChildByLocalName(wsp, "txbx"), "txbxContent") ?? undefined;
};

/**
 * The captured group XML with each edited text box's content replaced, or
 * undefined when an edit names no text box in it.
 *
 * The replacement is written as a placeholder element and substituted after
 * the tree is serialized, so the new content is never reparsed and the rest
 * of the capture is written exactly as it was read.
 */
export const replaceGroupTextBoxContent = (
  rawXml: string,
  edits: readonly GroupTextBoxEdit[],
): string | undefined => {
  const root = parseXml(rawXml);
  const group = findGroup(root);
  if (!group) {
    return undefined;
  }
  const replacements = new Map<string, string>();
  for (const [index, edit] of edits.entries()) {
    const content = textBoxContentAt(group, edit.path);
    if (!content) {
      return undefined;
    }
    const placeholder = `folio-group-text-${index}-${fingerprint(rawXml)}`;
    content.elements = [{ type: "element", name: placeholder }];
    // `CT_TxbxContent` holds at least one block, so an emptied box keeps one
    // empty paragraph.
    replacements.set(`<${placeholder}/>`, edit.contentXml || "<w:p/>");
  }
  let xml = (root.elements ?? []).map((element) => captureVerbatimXml(element)).join("");
  for (const [placeholder, contentXml] of replacements) {
    xml = xml.replace(placeholder, () => contentXml);
  }
  return xml;
};
