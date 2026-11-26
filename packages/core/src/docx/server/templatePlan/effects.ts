/**
 * Text effects that show nothing.
 *
 * A run's properties can carry text effects of the `w14` extension (`w14:glow`,
 * `w14:shadow`, `w14:reflection`, `w14:textOutline`, `w14:textFill`,
 * `w14:scene3d`, `w14:props3d`). An effect whose values make it invisible
 * still changes how some consumers treat the text: a run with a 3-D scene,
 * even the default one, is drawn as a picture, so its text cannot be found,
 * copied or read aloud in the exported document. Each rule here names the
 * values that make one effect invisible; an effect that meets no rule is
 * left alone.
 */
import { getLocalName, type XmlElement } from "../../xmlParser";
import type { DocxTextEffect, DocxTextEffectOwner } from "./types";
import {
  elementChildren,
  integerOrNull,
  isW,
  rangeOf,
  type RawPart,
  spliceXml,
  wAttr,
  wChild,
  type XmlRange,
} from "./xml";

const WORD_2010_NAMESPACE =
  "http://schemas.microsoft.com/office/word/2010/wordml";

export type TextEffect = DocxTextEffect;
export type TextEffectOwner = DocxTextEffectOwner;

export const TEXT_EFFECTS: readonly TextEffect[] = [
  "glow",
  "shadow",
  "reflection",
  "textOutline",
  "textFill",
  "scene3d",
  "props3d",
];

export type InvisibleEffect = {
  effect: TextEffect;
  owner: TextEffectOwner;
  range: XmlRange;
};

const isEffect = (element: XmlElement, effect: string): boolean =>
  element.type === "element" &&
  element.namespaceUri === WORD_2010_NAMESPACE &&
  getLocalName(element.name) === effect;

const effectChild = (element: XmlElement, name: string): XmlElement | null =>
  elementChildren(element).find((child) => isEffect(child, name)) ?? null;

/** An attribute of an effect element by its local name, whatever prefix
 *  the part binds the namespace to. */
const effectAttr = (element: XmlElement, name: string): string | null => {
  for (const [key, value] of Object.entries(element.attributes ?? {})) {
    if (value !== undefined && (key === name || key.endsWith(`:${name}`))) {
      return String(value);
    }
  }
  return null;
};

/** Whether every named attribute is stated and zero. */
const allZero = (element: XmlElement, names: readonly string[]): boolean =>
  names.every((name) => Number(effectAttr(element, name) ?? Number.NaN) === 0);

/** Whether a named attribute is zero, or left to its default of zero. */
const zeroOrAbsent = (element: XmlElement, name: string): boolean => {
  const value = effectAttr(element, name);
  return value === null || Number(value) === 0;
};

const FULLY_TRANSPARENT = 100_000;

/**
 * Whether the colour an effect is drawn in is fully transparent. `alpha`
 * states transparency here, in thousandths of a percent: the whole of it
 * leaves nothing of the effect.
 */
const drawnTransparent = (element: XmlElement): boolean =>
  elementChildren(element).some((colour) => {
    const alpha =
      isEffect(colour, "srgbClr") || isEffect(colour, "schemeClr")
        ? effectChild(colour, "alpha")
        : null;
    return (
      alpha !== null && Number(effectAttr(alpha, "val")) === FULLY_TRANSPARENT
    );
  });

/** Blur, distance and both scales at zero: nothing is drawn, whatever the
 *  colour, the direction or the skew. */
const ZERO_SIZE = ["blurRad", "dist", "sx", "sy"] as const;

/** No depth, no contour and no bevel: nothing for a 3-D scene to light. */
const flatProperties = (props3d: XmlElement): boolean =>
  zeroOrAbsent(props3d, "extrusionH") &&
  zeroOrAbsent(props3d, "contourW") &&
  effectChild(props3d, "bevelT") === null &&
  effectChild(props3d, "bevelB") === null;

/** The camera every flat text has: straight on, not turned. */
const defaultCamera = (scene3d: XmlElement): boolean => {
  const camera = effectChild(scene3d, "camera");
  const rotation = camera === null ? null : effectChild(camera, "rot");
  return (
    camera !== null &&
    effectAttr(camera, "prst") === "orthographicFront" &&
    (rotation === null || allZero(rotation, ["lat", "lon", "rev"]))
  );
};

const HEX_COLOUR = /^[0-9a-f]{6}$/iu;

/** A solid fill in exactly the colour the run states for its text. */
const fillIsRunColour = (
  textFill: XmlElement,
  runProperties: XmlElement,
): boolean => {
  const fills = elementChildren(textFill);
  const solid = fills.length === 1 ? effectChild(textFill, "solidFill") : null;
  const colours = solid === null ? [] : elementChildren(solid);
  const colour =
    colours.length === 1 ? effectChild(solid ?? textFill, "srgbClr") : null;
  const stated = wChild(runProperties, "color");
  const filled = colour === null ? null : effectAttr(colour, "val");
  const run = wAttr(stated, "val");
  return (
    colour !== null &&
    elementChildren(colour).length === 0 &&
    filled !== null &&
    run !== null &&
    wAttr(stated, "themeColor") === null &&
    HEX_COLOUR.test(filled) &&
    filled.toLowerCase() === run.toLowerCase()
  );
};

/**
 * Whether an effect of `runProperties` shows nothing. One rule per effect:
 *
 * - `glow`: radius zero, or a fully transparent colour.
 * - `shadow`: blur, distance and both scales zero, or a fully transparent
 *   colour. A shadow of full size straight under the text is not one: its
 *   edge shows around the glyphs.
 * - `reflection`: blur, distance and both scales zero.
 * - `textOutline`: no fill, a width of zero, or a fully transparent solid
 *   fill.
 * - `textFill`: a plain solid fill in the colour the run states. A
 *   transparent fill is the opposite of nothing: it hides the text.
 * - `props3d`: no extrusion, no contour, no bevel.
 * - `scene3d`: the default camera, not turned, and nothing 3-D beside it.
 */
const showsNothing = (
  effect: TextEffect,
  element: XmlElement,
  runProperties: XmlElement,
): boolean => {
  switch (effect) {
    case "glow":
      return allZero(element, ["rad"]) || drawnTransparent(element);
    case "shadow":
      return allZero(element, ZERO_SIZE) || drawnTransparent(element);
    case "reflection":
      return allZero(element, ZERO_SIZE);
    case "textOutline": {
      const solid = effectChild(element, "solidFill");
      return (
        allZero(element, ["w"]) ||
        (solid !== null && drawnTransparent(solid)) ||
        (effectChild(element, "noFill") !== null &&
          solid === null &&
          effectChild(element, "gradFill") === null)
      );
    }
    case "textFill":
      return fillIsRunColour(element, runProperties);
    case "props3d":
      return flatProperties(element);
    case "scene3d": {
      const props3d = effectChild(runProperties, "props3d");
      return (
        defaultCamera(element) && (props3d === null || flatProperties(props3d))
      );
    }
    default:
      return effect satisfies never;
  }
};

const ownerOf = (
  element: XmlElement,
  owner: TextEffectOwner,
): TextEffectOwner => {
  if (isW(element, "abstractNum")) {
    return {
      kind: "level",
      abstractNumId: integerOrNull(wAttr(element, "abstractNumId")),
      ilvl: null,
    };
  }
  if (isW(element, "lvl")) {
    return {
      kind: "level",
      abstractNumId: owner.kind === "level" ? owner.abstractNumId : null,
      ilvl: integerOrNull(wAttr(element, "ilvl")),
    };
  }
  if (isW(element, "style")) {
    return { kind: "style", styleId: wAttr(element, "styleId") };
  }
  return isW(element, "docDefaults") ? { kind: "defaults" } : owner;
};

/** Every text effect of a part that shows nothing, in document order. */
export const findInvisibleEffects = ({ root }: RawPart): InvisibleEffect[] => {
  const found: InvisibleEffect[] = [];
  const stack: { element: XmlElement; owner: TextEffectOwner }[] = [
    { element: root, owner: { kind: "text" } },
  ];
  for (let next = stack.pop(); next !== undefined; next = stack.pop()) {
    const { element } = next;
    const owner = ownerOf(element, next.owner);
    const children = elementChildren(element);
    if (isW(element, "rPr")) {
      for (const effect of TEXT_EFFECTS) {
        const stated = effectChild(element, effect);
        const range = stated === null ? null : rangeOf(stated);
        if (
          stated !== null &&
          range !== null &&
          showsNothing(effect, stated, element)
        ) {
          found.push({ effect, owner, range });
        }
      }
    }
    for (let index = children.length - 1; index >= 0; index -= 1) {
      const child = children[index];
      if (child !== undefined) {
        stack.push({ element: child, owner });
      }
    }
  }
  return found.toSorted((left, right) => left.range.start - right.range.start);
};

/**
 * `part` without its text effects that show nothing; every other byte
 * stays. `xml` is the part as it stands when nothing qualifies. Null when
 * the part cannot be rewritten.
 */
export const removeInvisibleEffects = (
  part: RawPart,
): { xml: string; removed: InvisibleEffect[] } | null => {
  const removed = findInvisibleEffects(part);
  if (removed.length === 0) {
    return { xml: part.xml, removed };
  }
  const xml = spliceXml(
    part.xml,
    removed.map(({ range }) => ({ ...range, newXml: "" })),
  );
  return xml === null ? null : { xml, removed };
};

/** How many of each effect a part lost, for a report. */
export const countEffects = (
  part: string,
  effects: readonly InvisibleEffect[],
): { part: string; effect: TextEffect; count: number }[] =>
  TEXT_EFFECTS.map((effect) => ({
    part,
    effect,
    count: effects.filter((found) => found.effect === effect).length,
  })).filter(({ count }) => count > 0);
