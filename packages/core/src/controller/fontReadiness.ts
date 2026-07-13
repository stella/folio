/**
 * Font-readiness helpers for the initial layout pass.
 *
 * Framework-neutral so both adapters share one implementation: collecting the
 * font faces a document needs (from its model + ProseMirror content), gating
 * the first layout on those faces having loaded, and reporting whether the
 * browser's font set is settled. Browser globals (`document.fonts`, `window`)
 * are only touched at call time, so importing this module in a non-browser host
 * is safe (mirrors the `browserClock` precedent in `layoutScheduler.ts`).
 */

import type { Mark, Node as PMNode } from "prosemirror-model";
import type { EditorState } from "prosemirror-state";

import { expectFontFamilyMarkAttrs } from "../prosemirror/attrs";
import type { Document, TextFormatting } from "../types/document";

export function getDocumentFontSet(): FontFaceSet | null {
  if (typeof document === "undefined" || !("fonts" in document)) {
    return null;
  }
  return document.fonts;
}

export function documentFontsAreLoaded(): boolean {
  const fontSet = getDocumentFontSet();
  return !fontSet || fontSet.status === "loaded";
}

const INITIAL_LAYOUT_FONT_TIMEOUT_MS = 2000;
const DEFAULT_LAYOUT_FONT_FAMILY = "Calibri";
const OFFICE_FONT_FAMILY_MAP: Record<string, string> = {
  Arial: "Arimo",
  Calibri: "Carlito",
  Cambria: "Caladea",
  "Times New Roman": "Tinos",
  "Courier New": "Cousine",
};
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

export function waitForInitialLayoutFonts(
  documentModel: Document | null,
  pmDoc: EditorState["doc"],
): Promise<boolean> {
  const fontSet = getDocumentFontSet();
  if (!fontSet) {
    return Promise.resolve(true);
  }

  const loadChecks: string[] = [];
  for (const face of collectInitialLayoutFontFaces(documentModel, pmDoc)) {
    loadChecks.push(`${face.style} ${face.weight} 16px "${escapeCssFontFamily(face.family)}"`);
  }

  const loadFonts = Promise.allSettled(loadChecks.map((check) => fontSet.load(check)))
    .then(() => fontSet.ready)
    .then(() => true);
  return Promise.race([
    loadFonts,
    new Promise<boolean>((resolve) => {
      globalThis.setTimeout(() => resolve(false), INITIAL_LAYOUT_FONT_TIMEOUT_MS);
    }),
  ]);
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
  addLayoutFontFamilyFace(faces, DEFAULT_LAYOUT_FONT_FAMILY, REGULAR_LAYOUT_FONT_DESCRIPTOR);

  for (const family of documentModel?.requiredFonts ?? []) {
    addLayoutFontFamilyFace(faces, family, REGULAR_LAYOUT_FONT_DESCRIPTOR);
  }

  addLayoutFontFamilyFace(
    faces,
    documentModel?.package.theme?.fontScheme?.majorFont?.latin,
    REGULAR_LAYOUT_FONT_DESCRIPTOR,
  );
  addLayoutFontFamilyFace(
    faces,
    documentModel?.package.theme?.fontScheme?.minorFont?.latin,
    REGULAR_LAYOUT_FONT_DESCRIPTOR,
  );
  addTextFormattingFontFaces(faces, documentModel?.package.styles?.docDefaults?.rPr);
  for (const style of documentModel?.package.styles?.styles ?? []) {
    addTextFormattingFontFaces(faces, style.rPr);
  }

  collectProseMirrorFontFaces(faces, pmDoc, undefined);

  return Array.from(faces.values());
}

function addTextFormattingFontFaces(
  faces: Map<string, LayoutFontFace>,
  formatting: TextFormatting | undefined,
): void {
  addLayoutFontFamilyFace(
    faces,
    formatting?.fontFamily,
    layoutDescriptorFromFormatting(formatting),
  );
}

function collectProseMirrorFontFaces(
  faces: Map<string, LayoutFontFace>,
  node: PMNode,
  inheritedTextFormatting: TextFormatting | undefined,
): void {
  const paragraphDefaults = readParagraphDefaultTextFormatting(node);
  const textFormatting = paragraphDefaults ?? inheritedTextFormatting;
  if (paragraphDefaults) {
    addTextFormattingFontFaces(faces, paragraphDefaults);
  }

  if (node.attrs["listMarkerFontFamily"]) {
    addLayoutFontFamilyFace(
      faces,
      node.attrs["listMarkerFontFamily"],
      REGULAR_LAYOUT_FONT_DESCRIPTOR,
    );
  }

  if (node.isText) {
    const descriptor = layoutDescriptorFromFormattingAndMarks(textFormatting, node.marks);
    const markFontFamily = readFontFamilyMarkAttrs(node.marks);
    addLayoutFontFamilyFace(
      faces,
      markFontFamily ?? textFormatting?.fontFamily ?? DEFAULT_LAYOUT_FONT_FAMILY,
      descriptor,
    );
  }

  // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Node.forEach
  node.forEach((child) => {
    collectProseMirrorFontFaces(faces, child, textFormatting);
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
  return {
    style: formatting?.italic ? "italic" : "normal",
    weight: formatting?.bold ? 700 : 400,
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
): void {
  if (typeof value === "string") {
    addLayoutFontFamilyNameFace(faces, value, descriptor);
    return;
  }

  if (!value || typeof value !== "object") {
    return;
  }

  const fontFamily = value as { ascii?: unknown; hAnsi?: unknown };
  addLayoutFontFamilyFace(faces, fontFamily.ascii, descriptor);
  addLayoutFontFamilyFace(faces, fontFamily.hAnsi, descriptor);
}

function addLayoutFontFamilyNameFace(
  faces: Map<string, LayoutFontFace>,
  family: string,
  descriptor: Omit<LayoutFontFace, "family">,
): void {
  const normalized = family.trim();
  if (!normalized || CSS_GENERIC_FONT_FAMILIES.has(normalized)) {
    return;
  }

  addLayoutFontFace(faces, normalized, descriptor);
  const mappedFamily = OFFICE_FONT_FAMILY_MAP[normalized];
  if (mappedFamily) {
    addLayoutFontFace(faces, mappedFamily, descriptor);
  }
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
