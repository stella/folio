/**
 * Embedded-font extraction.
 *
 * Turns the obfuscated `word/fonts/*.odttf` binaries a `.docx` carries into
 * de-obfuscated OpenType/TrueType faces the renderer can register as
 * `@font-face`s, so documents whose fonts are embedded (rather than installed)
 * render in their authored fonts instead of a fallback.
 *
 * The core is pure and framework-free: {@link getEmbeddedFontFaces} works over
 * already-unzipped package parts; {@link extractEmbeddedFonts} is the buffer
 * convenience that unzips first. Neither touches the DOM. Round-trip
 * preservation is unaffected: the serializer copies `word/fonts/*` and
 * `fontTable.xml` from the original ZIP untouched, so this is a read-only view.
 *
 * Ported from eigenpal/docx-editor (see NOTICE.md).
 */

import { parseRelationships, resolveRelativePath } from "../docx/relsParser";
import { unzipDocx } from "../docx/unzip";
import type { RelationshipMap } from "../types";
import { generateHexId } from "../utils/hexId";
import { deobfuscateFont, isValidFontKey } from "./fontDeobfuscation";
import { parseEmbeddedFontTable, type EmbeddedFontFaceRef } from "./embeddedFontTable";

// A DOCX author controls both the embedded font bytes AND the `w:font
// w:name` it's declared under — including names that collide with a real
// installed/host font (e.g. "Arial", "Inter"). Registering an embedded face
// under its raw name on the page-global `document.fonts` would let one
// document's embedded bytes shadow that family for the WHOLE host page,
// including host-page chrome that never opted into the DOCX's fonts. Every
// face is therefore registered under a per-document SCOPED name instead —
// see `scopeEmbeddedFontFamily` — and the raw name is never registered
// anywhere; `buildEmbeddedFontFamilyMap` gives callers the original→scoped
// mapping so the editor's own rendering can still resolve runs to it.
const SCOPED_FAMILY_PREFIX = "folio-embedded";

/**
 * Build the per-document scoped family name an embedded face registers
 * under. `docNonce` is fresh per document load (see {@link getEmbeddedFontFaces}),
 * so the same original name in two different documents (or two loads of the
 * same document) never collides.
 */
export function scopeEmbeddedFontFamily(docNonce: string, originalFamily: string): string {
  return `${SCOPED_FAMILY_PREFIX}-${docNonce}-${originalFamily}`;
}

/** A single de-obfuscated embedded font face, ready to register as `@font-face`. */
export type EmbeddedFont = {
  /**
   * Family name to register the face under (`@font-face`/`FontFace`).
   * Always a per-document SCOPED name (see {@link scopeEmbeddedFontFamily}) —
   * never the raw `w:font w:name` from the DOCX. Use `originalFamily` (and
   * {@link buildEmbeddedFontFamilyMap}) to resolve document runs, which
   * still reference the font by its original DOCX name, to this family.
   */
  family: string;
  /** Original Word font name (`w:font w:name`), before scoping. */
  originalFamily: string;
  /** CSS `font-style` the face maps to (`embed*Italic` → `italic`). */
  style: "normal" | "italic";
  /** CSS `font-weight` the face maps to (`embedBold*` → `700`). */
  weight: 400 | 700;
  /** De-obfuscated OpenType/TrueType bytes (backed by a fresh, non-shared buffer). */
  bytes: Uint8Array<ArrayBuffer>;
  /** Whether the source face was subsetted (`w:subsetted`). */
  subsetted: boolean;
};

/** Already-unzipped package parts the pure extractor needs. */
export type EmbeddedFontParts = {
  /** Raw `word/fontTable.xml`. */
  fontTableXml: string | null | undefined;
  /** Raw `word/_rels/fontTable.xml.rels`. */
  fontTableRelsXml: string | null | undefined;
  /** Unzipped font binaries keyed by package path (e.g. `word/fonts/font1.odttf`). */
  fonts: ReadonlyMap<string, ArrayBuffer>;
};

type EmbedField = "embedRegular" | "embedBold" | "embedItalic" | "embedBoldItalic";

type EmbedKind = {
  ref: EmbedField;
  weight: 400 | 700;
  style: "normal" | "italic";
};

/** Pairs each embed field with the CSS weight/style it renders as. */
const EMBED_KINDS = [
  { ref: "embedRegular", weight: 400, style: "normal" },
  { ref: "embedBold", weight: 700, style: "normal" },
  { ref: "embedItalic", weight: 400, style: "italic" },
  { ref: "embedBoldItalic", weight: 700, style: "italic" },
] as const satisfies readonly EmbedKind[];

// The font table's relationships always live at this OPC part path; embed
// targets resolve relative to it (see `resolveRelativePath`).
const FONTTABLE_RELS_BASE_PATH = "word/_rels/fontTable.xml.rels";
const FONTTABLE_RELS_PATH = "word/_rels/fonttable.xml.rels";

/** Normalize a package path for comparison: drop leading slashes, lowercase. */
function normalizePackagePath(path: string): string {
  return path.replace(/^\/+/u, "").toLowerCase();
}

/**
 * Find the font binary for an already-resolved package path (e.g.
 * `word/fonts/font1.odttf`) in the unzipped font map, matched case- and
 * leading-slash-insensitively (ZIP entry casing/prefixes vary across producers).
 */
function lookupFontData(
  resolvedPath: string,
  fonts: ReadonlyMap<string, ArrayBuffer>,
): ArrayBuffer | undefined {
  const wanted = normalizePackagePath(resolvedPath);
  for (const [path, data] of fonts) {
    if (normalizePackagePath(path) === wanted) {
      return data;
    }
  }
  return undefined;
}

function resolveFace(
  originalFamily: string,
  ref: EmbeddedFontFaceRef,
  kind: EmbedKind,
  fonts: ReadonlyMap<string, ArrayBuffer>,
  rels: RelationshipMap,
  docNonce: string,
): EmbeddedFont | null {
  const target = rels.get(ref.relId)?.target;
  if (!target) {
    return null;
  }

  // Resolve the rel target against the font-table part base the same way the
  // rest of the package resolves rels, so `fonts/font1.odttf`, `../media/...`,
  // and leading-slash absolute targets all land on the right ZIP entry.
  const resolvedPath = resolveRelativePath(FONTTABLE_RELS_BASE_PATH, target);
  const raw = lookupFontData(resolvedPath, fonts);
  if (!raw) {
    return null;
  }

  const bytes = new Uint8Array(raw);
  const data =
    ref.fontKey && isValidFontKey(ref.fontKey)
      ? deobfuscateFont(bytes, ref.fontKey)
      : // No usable key: some producers embed a non-obfuscated face verbatim.
        bytes;

  return {
    family: scopeEmbeddedFontFamily(docNonce, originalFamily),
    originalFamily,
    style: kind.style,
    weight: kind.weight,
    bytes: data,
    subsetted: ref.subsetted ?? false,
  };
}

/**
 * Resolve and de-obfuscate every embedded font face declared in a font table.
 * Pure and DOM-free. Faces whose relationship or binary is missing are skipped,
 * so the result only contains faces that produced usable bytes.
 *
 * `docNonce` scopes every face's registered family to one document load (see
 * {@link scopeEmbeddedFontFamily}); defaults to a fresh random id so callers
 * that don't care about a specific value still get per-call isolation. Pass
 * an explicit value for deterministic tests.
 */
export function getEmbeddedFontFaces(
  parts: EmbeddedFontParts,
  docNonce: string = generateHexId(),
): EmbeddedFont[] {
  const entries = parseEmbeddedFontTable(parts.fontTableXml);
  if (entries.length === 0) {
    return [];
  }
  if (!parts.fontTableRelsXml || parts.fonts.size === 0) {
    return [];
  }

  const rels = parseRelationships(parts.fontTableRelsXml);
  const faces: EmbeddedFont[] = [];
  for (const entry of entries) {
    for (const kind of EMBED_KINDS) {
      const ref = entry[kind.ref];
      if (!ref) {
        continue;
      }
      const face = resolveFace(entry.name, ref, kind, parts.fonts, rels, docNonce);
      if (face) {
        faces.push(face);
      }
    }
  }
  return faces;
}

/**
 * Build the original DOCX font name → scoped registered family map for a set
 * of resolved embedded faces. Callers thread this into font resolution (see
 * `utils/fontResolver.ts`'s `setEmbeddedFontFamilyMap`) so the document's own
 * runs — which still carry the original name — keep resolving to the scoped
 * face that was actually registered on `document.fonts`.
 */
export function buildEmbeddedFontFamilyMap(faces: readonly EmbeddedFont[]): Map<string, string> {
  const map = new Map<string, string>();
  for (const face of faces) {
    map.set(face.originalFamily, face.family);
  }
  return map;
}

function findFontTableRels(allXml: ReadonlyMap<string, string>): string | undefined {
  for (const [path, xml] of allXml) {
    if (normalizePackagePath(path) === FONTTABLE_RELS_PATH) {
      return xml;
    }
  }
  return undefined;
}

/**
 * Extract every embedded font face from a raw `.docx` buffer. Unzips the
 * package, then resolves + de-obfuscates the faces via
 * {@link getEmbeddedFontFaces}. Returns an empty list for documents with no
 * embedded fonts. See {@link getEmbeddedFontFaces} for `docNonce`.
 */
export async function extractEmbeddedFonts(
  buffer: ArrayBuffer,
  docNonce: string = generateHexId(),
): Promise<EmbeddedFont[]> {
  const raw = await unzipDocx(buffer);
  return getEmbeddedFontFaces(
    {
      fontTableXml: raw.fontTableXml,
      fontTableRelsXml: findFontTableRels(raw.allXml),
      fonts: raw.fonts,
    },
    docNonce,
  );
}
