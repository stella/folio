/**
 * `word/fontTable.xml`: every font a document declares, and what it declares.
 *
 * The part is copied byte for byte by a repack, so nothing here runs on an
 * ordinary save. It runs on the paths that rebuild a package from the model —
 * a package folio authors, a style set carried into a new document, and the
 * survival census's part leg — and on those paths whatever this reader drops
 * is gone. So the walk goes through the shared child dispatcher: the handler
 * map is total over the children `CT_Font` declares, and anything else lands
 * in the ordered sink beside the font's own fields.
 */

import type { EmbeddedFontRef, FontCharset, FontInfo, FontTable } from "../types/document";

import { attributeRemainder, DERIVED_PART_ROOT_ATTRIBUTES } from "./attributeRemainder";
import { CAPTURE, type ChildHandlers, dispatchChildrenWithContext } from "./containerChildren";
import { getAttribute, parseXmlDocument } from "./xmlParser";
import type { XmlElement } from "./xmlParser";

const FONT_FAMILIES = ["decorative", "modern", "roman", "script", "swiss", "auto"] as const;
const FONT_PITCHES = ["default", "fixed", "variable"] as const;

/** `w:font` reads its name; everything else it carries is the remainder's. */
const MODELLED_FONT_ATTRIBUTES: ReadonlySet<string> = new Set(["name"]);

export const parseFontTable = (xml: string | null | undefined): FontTable | undefined => {
  if (!xml) {
    return undefined;
  }
  const root = parseXmlDocument(xml);
  if (!root) {
    return undefined;
  }
  const fonts: FontInfo[] = [];
  const preserved = dispatchChildrenWithContext({
    element: root,
    container: "w:fonts",
    capturePosition: () => fonts.length,
    handlers: FONT_TABLE_HANDLERS,
    context: fonts,
  });
  const preservedAttributes = attributeRemainder({
    element: root,
    modelled: DERIVED_PART_ROOT_ATTRIBUTES,
  });
  if (fonts.length === 0 && preserved === undefined && preservedAttributes === undefined) {
    return undefined;
  }
  return {
    fonts,
    ...(preserved === undefined ? {} : { preserved }),
    ...(preservedAttributes === undefined ? {} : { preservedAttributes }),
  };
};

const parseFont = (element: XmlElement): FontInfo | undefined => {
  const name = getAttribute(element, "w", "name");
  if (!name) {
    return undefined;
  }
  const walk: FontWalk = { font: { name }, modelled: 0 };
  const preserved = dispatchChildrenWithContext({
    element,
    container: "w:font",
    capturePosition: () => walk.modelled,
    handlers: FONT_HANDLERS,
    context: walk,
  });
  const { font } = walk;
  if (preserved !== undefined) {
    font.preserved = preserved;
  }
  const preservedAttributes = attributeRemainder({
    element,
    modelled: MODELLED_FONT_ATTRIBUTES,
  });
  if (preservedAttributes !== undefined) {
    font.preservedAttributes = preservedAttributes;
  }
  return font;
};

const value = (element: XmlElement): string | undefined =>
  getAttribute(element, "w", "val") ?? undefined;

/**
 * An empty `w:charset` is not an absent one.
 *
 * Both of `CT_Charset`'s attributes are optional, so a bare `<w:charset/>` is
 * legal and says the font uses the default code page. Reading it as nothing
 * would write it as nothing, and the document would come back saying something
 * else: the record is what carries the element's presence.
 */
const charsetOf = (element: XmlElement): FontCharset => {
  const val = value(element);
  const characterSet = getAttribute(element, "w", "characterSet") ?? undefined;
  return {
    ...(val === undefined ? {} : { val }),
    ...(characterSet === undefined ? {} : { characterSet }),
  };
};

/**
 * An embedded face with no relationship points at no binary.
 *
 * Writing the element back without one would declare a face nothing can
 * resolve, so the reference is only a reference when it has an `r:id`.
 */
const embedOf = (element: XmlElement): EmbeddedFontRef | undefined => {
  const id = getAttribute(element, "r", "id");
  if (!id) {
    return undefined;
  }
  const fontKey = getAttribute(element, "w", "fontKey") ?? undefined;
  const subsetted = getAttribute(element, "w", "subsetted");
  return {
    id,
    ...(fontKey === undefined ? {} : { fontKey }),
    ...(subsetted === null || subsetted === undefined
      ? {}
      : { subsetted: subsetted === "true" || subsetted === "1" || subsetted === "on" }),
  };
};

const SIGNATURE_ATTRIBUTES = ["usb0", "usb1", "usb2", "usb3", "csb0", "csb1"] as const;

const signatureOf = (element: XmlElement): FontInfo["sig"] | undefined => {
  const signature: NonNullable<FontInfo["sig"]> = {};
  for (const attribute of SIGNATURE_ATTRIBUTES) {
    const carried = getAttribute(element, "w", attribute);
    if (carried !== null && carried !== undefined) {
      signature[attribute] = carried;
    }
  }
  return Object.keys(signature).length === 0 ? undefined : signature;
};

const narrow = <Member extends string>(
  carried: string | undefined,
  members: readonly Member[],
): Member | undefined => members.find((member) => member === carried);

const FONT_TABLE_HANDLERS = {
  font: (element, fonts) => {
    const font = parseFont(element);
    if (font) {
      fonts.push(font);
    }
  },
} as const satisfies ChildHandlers<"w:fonts", FontInfo[]>;

/**
 * `CT_Font` is a sequence and the serializer writes its fields back in that
 * order, so the sink's index is a count of the fields read so far and a
 * capture lands between the same two neighbours it sat between.
 */
type FontWalk = { font: FontInfo; modelled: number };

const read = <Key extends keyof FontInfo>(
  walk: FontWalk,
  key: Key,
  carried: FontInfo[Key] | undefined,
): void => {
  if (carried === undefined) {
    return;
  }
  walk.font[key] = carried;
  walk.modelled += 1;
};

const FONT_HANDLERS = {
  altName: (child, walk) => read(walk, "altName", value(child)),
  panose1: (child, walk) => read(walk, "panose1", value(child)),
  charset: (child, walk) => read(walk, "charset", charsetOf(child)),
  family: (child, walk) => read(walk, "family", narrow(value(child), FONT_FAMILIES)),
  // Word's own flag for a face that is not TrueType. folio has no model
  // for it and no reader asks, so the bytes go back where they were.
  notTrueType: CAPTURE,
  pitch: (child, walk) => read(walk, "pitch", narrow(value(child), FONT_PITCHES)),
  sig: (child, walk) => read(walk, "sig", signatureOf(child)),
  embedRegular: (child, walk) => read(walk, "embedRegular", embedOf(child)),
  embedBold: (child, walk) => read(walk, "embedBold", embedOf(child)),
  embedItalic: (child, walk) => read(walk, "embedItalic", embedOf(child)),
  embedBoldItalic: (child, walk) => read(walk, "embedBoldItalic", embedOf(child)),
} as const satisfies ChildHandlers<"w:font", FontWalk>;
