/**
 * Font readiness for layout.
 *
 * Framework-neutral so both adapters share one implementation: collecting the
 * font faces and the characters a document needs (from its model + ProseMirror
 * content), gating the first layout on those faces having loaded, naming the
 * font set a layout was measured in, and re-running layout when a face that
 * can change a measured width loads afterwards. Browser globals
 * (`document.fonts`, `window`) are only touched at call time, so importing this
 * module in a non-browser host is safe (mirrors the `browserClock` precedent in
 * `layoutScheduler.ts`).
 */

import type { Mark, Node as PMNode } from "prosemirror-model";
import type { EditorState } from "prosemirror-state";

import {
  buildFontAlternates,
  getFontAlternate,
  type FontAlternates,
} from "../fonts/fontAlternates";
import { blockPlainText } from "../docx/blockPlainText";
import { expectFontFamilyMarkAttrs, expectParagraphAttrs } from "../prosemirror/attrs";
import { DEFAULT_FONT_FAMILY, isLayoutFontFamily } from "../layout-engine/measure/measureHelpers";
import { parseFontFamilyList, resolveFontFamily } from "../utils/fontResolver";
import type { Document, TextFormatting } from "../types/document";

/** The parts of a `FontFace` layout readiness reads. */
export type LayoutFontSetFace = Pick<FontFace, "family" | "status">;

type LayoutFontSetListener = (event: { fontfaces: readonly LayoutFontSetFace[] }) => void;

/**
 * The parts of a `FontFaceSet` layout readiness uses; `document.fonts` in a
 * browser, a scripted stand-in in tests.
 */
export type LayoutFontSet = Iterable<LayoutFontSetFace> & {
  load: (font: string, text?: string) => Promise<unknown>;
  readonly ready: Promise<unknown>;
  addEventListener: (type: "loadingdone", listener: LayoutFontSetListener) => void;
  removeEventListener: (type: "loadingdone", listener: LayoutFontSetListener) => void;
};

export function getDocumentFontSet(): FontFaceSet | null {
  if (typeof document === "undefined" || !("fonts" in document)) {
    return null;
  }
  return document.fonts;
}

/**
 * Names the faces available to measurement: which faces of the families
 * measured stacks name have loaded. Equal signatures mean every measured glyph
 * resolves to the same face, so a measurement taken under one is valid under
 * the other; a host UI face loading leaves it unchanged.
 *
 * Counted over loaded faces only (a face never unloads, and a failed one adds
 * no glyphs) with a per-face identity sum, so a removal replaced by an addition
 * still changes it: a new face's id exceeds every earlier one. The family set
 * only grows, and a family joins it when first measured, so read it after
 * measuring to name what a layout measured in. Without a font set (a server or
 * headless host) there is nothing to load and the signature is constant.
 */
export function readFontSetSignature(fontSet: LayoutFontSet | null = getDocumentFontSet()): string {
  if (!fontSet) {
    return NO_FONT_SET_SIGNATURE;
  }
  let loadedFaces = 0;
  let loadedFaceIdSum = 0;
  for (const face of fontSet) {
    if (face.status !== "loaded" || !isLayoutFontFamily(face.family)) {
      continue;
    }
    loadedFaces += 1;
    loadedFaceIdSum += fontFaceId(face);
  }
  return `${loadedFaces}:${loadedFaceIdSum}`;
}

const NO_FONT_SET_SIGNATURE = "none";
const fontFaceIds = new WeakMap<LayoutFontSetFace, number>();
let nextFontFaceId = 1;

const fontFaceId = (face: LayoutFontSetFace): number => {
  const known = fontFaceIds.get(face);
  if (known !== undefined) {
    return known;
  }
  const id = nextFontFaceId;
  nextFontFaceId += 1;
  fontFaceIds.set(face, id);
  return id;
};

export type WatchLayoutFontLoadsOptions = {
  /**
   * The font-set signature the committed layout was measured in, or `null`
   * before one is committed (the first layout measures whatever has loaded).
   */
  measuredFontSet: () => string | null;
  relayout: () => void;
  fontSet?: LayoutFontSet | null;
};

/**
 * Re-run layout whenever a face that can change a measured width finishes
 * loading after the committed layout measured without it.
 *
 * Tied to what loaded, never to a time window: fontsource splits each face
 * into `unicode-range` subsets the browser fetches when painted text first
 * needs them, so a subset can land at any moment after the first layout. A
 * load the committed layout already saw (its event can arrive after the layout
 * ran) leaves the signature unchanged and relays nothing, as does a face of a
 * family no measured stack names (a host UI font). Returns the unsubscribe
 * function.
 */
export function watchLayoutFontLoads({
  measuredFontSet,
  relayout,
  fontSet = getDocumentFontSet(),
}: WatchLayoutFontLoadsOptions): () => void {
  if (!fontSet) {
    return () => undefined;
  }
  const handleLoadingDone: LayoutFontSetListener = () => {
    const measured = measuredFontSet();
    if (measured === null || measured === readFontSetSignature(fontSet)) {
      return;
    }
    relayout();
  };
  fontSet.addEventListener("loadingdone", handleLoadingDone);
  return () => fontSet.removeEventListener("loadingdone", handleLoadingDone);
}

const INITIAL_LAYOUT_FONT_TIMEOUT_MS = 2000;
const CSS_GENERIC_FONT_FAMILIES = new Set([
  "serif",
  "sans-serif",
  "monospace",
  "cursive",
  "fantasy",
  "system-ui",
]);
const LAYOUT_FONT_DESCRIPTORS = [
  { style: "normal", weight: 400 },
  { style: "italic", weight: 400 },
  { style: "normal", weight: 700 },
  { style: "italic", weight: 700 },
] as const;
const REGULAR_LAYOUT_FONT_DESCRIPTOR = LAYOUT_FONT_DESCRIPTORS[0];

export type LayoutFontFace = {
  family: string;
  style: (typeof LAYOUT_FONT_DESCRIPTORS)[number]["style"];
  weight: (typeof LAYOUT_FONT_DESCRIPTORS)[number]["weight"];
};

/**
 * Resolve once every face the first layout will measure has loaded (`true`),
 * or after a timeout (`false`), so the first layout does not measure fallbacks.
 *
 * Each face is loaded for the document's own characters. A bundled face is a
 * set of `unicode-range` subsets, and `FontFaceSet.load` fetches only the
 * subsets covering the text it is given; without text it takes a single space
 * and loads only the Latin subset, so Czech, Polish, Greek and Cyrillic were
 * measured in a fallback.
 */
export function waitForInitialLayoutFonts(
  documentModel: Document | null,
  pmDoc: EditorState["doc"],
  fontSet: LayoutFontSet | null = getDocumentFontSet(),
): Promise<boolean> {
  if (!fontSet) {
    return Promise.resolve(true);
  }

  const text = collectLayoutText(documentModel, pmDoc);
  const loadChecks: string[] = [];
  for (const face of collectInitialLayoutFontFaces(documentModel, pmDoc)) {
    loadChecks.push(`${face.style} ${face.weight} 16px "${escapeCssFontFamily(face.family)}"`);
  }

  let timeout: ReturnType<typeof globalThis.setTimeout> | undefined;
  const loadFonts = Promise.allSettled(loadChecks.map((check) => fontSet.load(check, text)))
    .then(() => fontSet.ready)
    .then(() => true);
  return Promise.race([
    loadFonts,
    new Promise<boolean>((resolve) => {
      timeout = globalThis.setTimeout(() => resolve(false), INITIAL_LAYOUT_FONT_TIMEOUT_MS);
    }),
  ]).finally(() => globalThis.clearTimeout(timeout));
}

/**
 * Every distinct character the layout will measure, as one string: the body
 * text (tracked deletions included, since All Markup paints them), list
 * markers, headers, footers and notes. Always carries a space, which every
 * line measures.
 */
export function collectLayoutText(
  documentModel: Document | null,
  pmDoc: EditorState["doc"],
): string {
  const characters = new Set<string>([" "]);
  const addText = (text: string) => {
    for (const character of text) {
      if (character >= " ") {
        characters.add(character);
      }
    }
  };

  pmDoc.descendants((node) => {
    if (node.isText) {
      addText(node.text ?? "");
    } else if (node.type.name === "paragraph") {
      addText(expectParagraphAttrs(node).listMarker ?? "");
    }
    return true;
  });

  const documentPackage = documentModel?.package;
  for (const story of [
    ...(documentPackage?.headers?.values() ?? []),
    ...(documentPackage?.footers?.values() ?? []),
    ...(documentPackage?.footnotes ?? []),
    ...(documentPackage?.endnotes ?? []),
  ]) {
    addText(blockPlainText(story.content));
  }

  return Array.from(characters).join("");
}

export function collectInitialLayoutFontFamilies(
  documentModel: Document | null,
  pmDoc: EditorState["doc"],
): Set<string> {
  return new Set(collectInitialLayoutFontFaces(documentModel, pmDoc).map(({ family }) => family));
}

export function collectInitialLayoutFontFaces(
  documentModel: Document | null,
  pmDoc: EditorState["doc"],
): LayoutFontFace[] {
  const faces = new Map<string, LayoutFontFace>();
  const fontAlternates = buildFontAlternates(documentModel?.package.fontTable);
  addLayoutFontFamilyFace(
    faces,
    DEFAULT_FONT_FAMILY,
    REGULAR_LAYOUT_FONT_DESCRIPTOR,
    fontAlternates,
  );

  for (const family of documentModel?.requiredFonts ?? []) {
    addLayoutFontFamilyFace(faces, family, REGULAR_LAYOUT_FONT_DESCRIPTOR, fontAlternates);
  }

  addLayoutFontFamilyFace(
    faces,
    documentModel?.package.theme?.fontScheme?.majorFont?.latin,
    REGULAR_LAYOUT_FONT_DESCRIPTOR,
    fontAlternates,
  );
  addLayoutFontFamilyFace(
    faces,
    documentModel?.package.theme?.fontScheme?.minorFont?.latin,
    REGULAR_LAYOUT_FONT_DESCRIPTOR,
    fontAlternates,
  );
  addTextFormattingFontFaces(
    faces,
    documentModel?.package.styles?.docDefaults?.rPr,
    fontAlternates,
  );
  for (const style of documentModel?.package.styles?.styles ?? []) {
    addTextFormattingFontFaces(faces, style.rPr, fontAlternates);
  }

  collectProseMirrorFontFaces(faces, pmDoc, undefined, fontAlternates);

  return Array.from(faces.values());
}

function addTextFormattingFontFaces(
  faces: Map<string, LayoutFontFace>,
  formatting: TextFormatting | undefined,
  fontAlternates: FontAlternates,
): void {
  const standardDescriptor = layoutDescriptorFromFormatting(formatting);
  const complexScriptDescriptor = layoutDescriptorFromEmphasis(
    formatting?.boldCs ?? formatting?.bold,
    formatting?.italicCs ?? formatting?.italic,
  );
  addLayoutFontFamilyFace(
    faces,
    formatting?.fontFamily,
    standardDescriptor,
    fontAlternates,
    complexScriptDescriptor,
  );
}

function collectProseMirrorFontFaces(
  faces: Map<string, LayoutFontFace>,
  node: PMNode,
  inheritedTextFormatting: TextFormatting | undefined,
  fontAlternates: FontAlternates,
): void {
  const paragraphDefaults = readParagraphDefaultTextFormatting(node);
  const textFormatting = paragraphDefaults ?? inheritedTextFormatting;
  if (paragraphDefaults) {
    addTextFormattingFontFaces(faces, paragraphDefaults, fontAlternates);
  }

  if (node.type.name === "paragraph") {
    addTextFormattingFontFaces(
      faces,
      expectParagraphAttrs(node).listMarkerFormatting,
      fontAlternates,
    );
  }

  if (node.isText) {
    const descriptor = layoutDescriptorFromFormattingAndMarks(textFormatting, node.marks);
    const markFontFamily = readFontFamilyMarkAttrs(node.marks);
    addLayoutFontFamilyFace(
      faces,
      markFontFamily ?? textFormatting?.fontFamily ?? DEFAULT_FONT_FAMILY,
      descriptor,
      fontAlternates,
    );
  }

  // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Node.forEach
  node.forEach((child) => {
    collectProseMirrorFontFaces(faces, child, textFormatting, fontAlternates);
  });
}

function readParagraphDefaultTextFormatting(node: PMNode): TextFormatting | undefined {
  const value = node.attrs["defaultTextFormatting"];
  if (!value || typeof value !== "object") {
    return undefined;
  }
  return value as TextFormatting;
}

function readFontFamilyMarkAttrs(marks: readonly Mark[]): unknown {
  for (const mark of marks) {
    if (mark.type.name === "fontFamily") {
      return expectFontFamilyMarkAttrs(mark);
    }
  }
  return undefined;
}

function layoutDescriptorFromFormatting(
  formatting: Pick<TextFormatting, "bold" | "italic"> | undefined,
): Omit<LayoutFontFace, "family"> {
  return layoutDescriptorFromEmphasis(formatting?.bold, formatting?.italic);
}

function layoutDescriptorFromEmphasis(
  bold: boolean | undefined,
  italic: boolean | undefined,
): Omit<LayoutFontFace, "family"> {
  return {
    style: italic ? "italic" : "normal",
    weight: bold ? 700 : 400,
  };
}

function layoutDescriptorFromFormattingAndMarks(
  formatting: Pick<TextFormatting, "bold" | "italic"> | undefined,
  marks: readonly Mark[],
): Omit<LayoutFontFace, "family"> {
  let bold = formatting?.bold === true;
  let italic = formatting?.italic === true;

  for (const mark of marks) {
    if (mark.type.name === "bold") {
      bold = true;
    }
    if (mark.type.name === "italic") {
      italic = true;
    }
  }

  return {
    style: italic ? "italic" : "normal",
    weight: bold ? 700 : 400,
  };
}

function addLayoutFontFamilyFace(
  faces: Map<string, LayoutFontFace>,
  value: unknown,
  descriptor: Omit<LayoutFontFace, "family">,
  fontAlternates: FontAlternates,
  complexScriptDescriptor = descriptor,
): void {
  if (typeof value === "string") {
    addLayoutFontFamilyNameFace(faces, value, descriptor, fontAlternates);
    return;
  }

  if (!value || typeof value !== "object") {
    return;
  }

  // Every slot, not just the Latin ones. `w:cs` carries the complex-script face
  // (Word writes Arabic and Hebrew there) and `w:eastAsia` the CJK face, so
  // collecting only ascii/hAnsi left the first layout measuring a fallback for
  // exactly the scripts whose advances differ most from it, then repainting in
  // the real face once it loaded.
  const fontFamily = value as {
    ascii?: unknown;
    hAnsi?: unknown;
    cs?: unknown;
    eastAsia?: unknown;
  };
  addLayoutFontFamilyFace(faces, fontFamily.ascii, descriptor, fontAlternates);
  addLayoutFontFamilyFace(faces, fontFamily.hAnsi, descriptor, fontAlternates);
  addLayoutFontFamilyFace(faces, fontFamily.cs, complexScriptDescriptor, fontAlternates);
  addLayoutFontFamilyFace(faces, fontFamily.eastAsia, descriptor, fontAlternates);
}

function addLayoutFontFamilyNameFace(
  faces: Map<string, LayoutFontFace>,
  family: string,
  descriptor: Omit<LayoutFontFace, "family">,
  fontAlternates: FontAlternates,
): void {
  const normalized = family.trim();
  if (!normalized || CSS_GENERIC_FONT_FAMILIES.has(normalized)) {
    return;
  }

  addLayoutFontFace(faces, normalized, descriptor);
  const alternate = getFontAlternate(normalized, fontAlternates);
  if (alternate) {
    addLayoutFontFace(faces, alternate, descriptor);
  }

  // Wait for the whole stack the renderer will actually use, not just the name
  // the document wrote. `resolveFontFamily` appends folio's bundled substitutes
  // and a script fallback, so an authored "Arial" run paints its Arabic in the
  // bundled Arabic face. Collecting only authored names meant the gate released
  // the first layout before that face had loaded, and measurement taken against
  // the pre-load fallback disagreed with what was ultimately painted.
  //
  // Derived rather than listed: a hand-kept table of substitutes would be a
  // second copy of the resolver's mapping, free to drift from it.
  for (const stackFamily of resolvedStackFamilies(normalized, alternate)) {
    addLayoutFontFace(faces, stackFamily, descriptor);
  }
}

/**
 * The concrete families in a resolved CSS font stack, generics dropped.
 *
 * Parsed from the stack rather than read from a map because the stack is what
 * the painter and the measurer put in `ctx.font` and `style.fontFamily`.
 */
function resolvedStackFamilies(family: string, alternate: string | undefined): string[] {
  const { cssFallback } = resolveFontFamily(family, alternate);
  return parseFontFamilyList(cssFallback).filter((name) => !CSS_GENERIC_FONT_FAMILIES.has(name));
}

function addLayoutFontFace(
  faces: Map<string, LayoutFontFace>,
  family: string,
  descriptor: Omit<LayoutFontFace, "family">,
): void {
  faces.set(`${family}|${descriptor.style}|${descriptor.weight}`, {
    family,
    ...descriptor,
  });
}

function escapeCssFontFamily(family: string): string {
  return family.replace(/\\/gu, "\\\\").replace(/"/gu, '\\"');
}
