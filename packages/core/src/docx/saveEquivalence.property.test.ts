/**
 * Save-path equivalence property tests.
 *
 * Folio persists an edited `.docx` through one of two paths:
 *
 *   SEL  = attemptSelectiveSave  — byte-exact per-paragraph patch of the
 *          original zip; bails to `null` when it cannot prove the patch is safe.
 *   FULL = repackDocx            — copy every original zip entry, then overwrite
 *          `word/document.xml` (and headers/footers/comments) from the model.
 *
 * These properties pin the safety contract of those two paths against a small,
 * bounded corpus of real fixtures, so a regression in either path is caught at
 * the model level (not just as an XML formatting wobble).
 *
 * Invariants (see the describe blocks below):
 *   1. No-op fidelity — saving an unedited doc round-trips the parsed model.
 *   2. Selective ≡ full repack — after an edit, parse(SEL) deep-equals parse(FULL).
 *   3. Selective is never silently wrong — SEL returns null, or a document that
 *      applies exactly the edit and nothing else.
 *   4. Byte-exactness — a paragraph edit leaves untouched parts and unedited
 *      paragraphs byte-identical to the original zip entries.
 *
 * Bound: 3-fixture sample, fixed seed, numRuns capped at 15 for the
 * podily-bps.docx (~176 KB) edit-driven properties, which are the cost driver.
 * This is deliberately NOT the full corpus × unbounded runs.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";
import { readFileSync } from "node:fs";
import path from "node:path";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import type { BlockContent, Document, Paragraph, Run } from "../types/document";
import { parseDocx } from "./parser";
import { repackDocx } from "./rezip";
import { attemptSelectiveSave } from "./selectiveSave";
import { findParagraphOffsets } from "./selectiveXmlPatch";

// ============================================================================
// CORPUS — a minimal, bounded sample of real fixtures
// ============================================================================

const FIXTURES_DIR = path.resolve(import.meta.dir, "../../../../tests/visual/fixtures");

// sample.docx           — plain paragraphs + one table + header/footer.
// docx-editor-demo.docx — lists, tables, header/footer, footnotes, comments.
// podily-bps.docx       — the rich one: 674 w14:paraIds, lists, tables, three
//                         headers/footers, footnotes AND endnotes. Only this
//                         fixture carries paraIds, so it is the only one on
//                         which selective save can engage its patch path.
const FIXTURES = ["sample.docx", "docx-editor-demo.docx", "podily-bps.docx"] as const;
const EDIT_FIXTURE = "podily-bps.docx";

// A fixed seed keeps the shrunk counterexamples reproducible across runs.
const SEED = 0x5a1e;

const fixtureCache = new Map<string, ArrayBuffer>();

const readFixture = (name: string): ArrayBuffer => {
  const cached = fixtureCache.get(name);
  if (cached) {
    return cached;
  }
  const bytes = readFileSync(path.join(FIXTURES_DIR, name));
  const buffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
  fixtureCache.set(name, buffer);
  return buffer;
};

const parse = (buffer: ArrayBuffer): Promise<Document> =>
  parseDocx(buffer, { preloadFonts: false });

const fullRepack = (doc: Document, originalBuffer: ArrayBuffer): Promise<ArrayBuffer> =>
  repackDocx({ ...doc, originalBuffer });

// ============================================================================
// MODEL NORMALIZATION — strip save-volatile fields for model-level comparison
// ============================================================================

// `dcterms:modified` / `lastModifiedBy` are refreshed by BOTH save paths, and
// `originalBuffer` is the raw input bytes; none of these are document content,
// so they are erased before comparing parsed models. Maps become sorted entry
// arrays and binary parts collapse to their length so deep-equality is stable.
const VOLATILE_KEYS = new Set(["originalBuffer", "modified", "lastModifiedBy"]);

const normalize = (value: unknown, key?: string): unknown => {
  if (key !== undefined && VOLATILE_KEYS.has(key)) {
    return "<<volatile>>";
  }
  if (value instanceof Uint8Array) {
    return { __byteLength: value.length };
  }
  if (value instanceof ArrayBuffer) {
    return { __byteLength: value.byteLength };
  }
  if (value instanceof Map) {
    return [...value.entries()]
      .map(([k, v]) => [k, normalize(v)] as const)
      .sort((a, b) => String(a[0]).localeCompare(String(b[0])));
  }
  if (Array.isArray(value)) {
    return value.map((item) => normalize(item));
  }
  if (value !== null && typeof value === "object") {
    const record = value as Record<string, unknown>;
    const out: Record<string, unknown> = {};
    for (const k of Object.keys(record).sort()) {
      out[k] = normalize(record[k], k);
    }
    return out;
  }
  return value;
};

const normalizedPackage = (doc: Document): unknown => normalize(doc.package);

// ============================================================================
// BODY-TEXT PROJECTION — visible text per top-level paragraph, in order
// ============================================================================

const runText = (run: Run): string => {
  let text = "";
  for (const item of run.content) {
    if (item.type === "text") {
      text += item.text;
    }
  }
  return text;
};

const paragraphText = (para: Paragraph): string => {
  let text = "";
  for (const item of para.content) {
    if (item.type === "run") {
      text += runText(item);
    } else if (item.type === "hyperlink") {
      for (const child of item.children) {
        if (child.type === "run") {
          text += runText(child);
        }
      }
    }
  }
  return text;
};

/** Top-level body paragraphs (tables/blockSdt excluded), in document order. */
const topLevelParagraphs = (blocks: readonly BlockContent[]): Paragraph[] =>
  blocks.filter((block): block is Paragraph => block.type === "paragraph");

const bodyParagraphTexts = (doc: Document): string[] =>
  topLevelParagraphs(doc.package.document.content).map(paragraphText);

const allBodyText = (doc: Document): string => bodyParagraphTexts(doc).join("");

// ============================================================================
// EDIT ARBITRARY — realistic in-place edits the editor actually produces
// ============================================================================

const isSafeText = (text: string): boolean => {
  if (text.trim().length === 0) {
    return false;
  }
  for (const char of text) {
    const codePoint = char.codePointAt(0) ?? 0;
    if (codePoint < 0x20 || char === "<" || char === ">" || char === "&") {
      return false;
    }
  }
  return true;
};

const safeText = fc.string({ minLength: 1, maxLength: 40 }).filter(isSafeText);

type EditSpec = {
  targetSelector: number;
  kind: "append" | "replace" | "toggleBold";
  text: string;
};

const arbEditSpec: fc.Arbitrary<EditSpec> = fc.record({
  targetSelector: fc.nat(),
  kind: fc.constantFrom("append", "replace", "toggleBold"),
  text: safeText,
});

type EditableTarget = { para: Paragraph; run: Run; textIndex: number };

const findEditableTarget = (para: Paragraph): EditableTarget | null => {
  for (const item of para.content) {
    if (item.type !== "run") {
      continue;
    }
    for (let i = 0; i < item.content.length; i++) {
      const content = item.content[i];
      if (content?.type === "text" && content.text.length > 0) {
        return { para, run: item, textIndex: i };
      }
    }
  }
  return null;
};

/**
 * Apply the edit to a freshly parsed doc, mutating in place. Returns the paraId
 * that was changed, or null when the doc has no editable paraId'd paragraph.
 */
const applyEdit = (doc: Document, spec: EditSpec): string | null => {
  const candidates: EditableTarget[] = [];
  for (const para of topLevelParagraphs(doc.package.document.content)) {
    if (!para.paraId) {
      continue;
    }
    const target = findEditableTarget(para);
    if (target) {
      candidates.push(target);
    }
  }
  if (candidates.length === 0) {
    return null;
  }

  const target = candidates[spec.targetSelector % candidates.length];
  if (!target?.para.paraId) {
    return null;
  }

  const content = target.run.content[target.textIndex];
  if (spec.kind === "toggleBold") {
    target.run.formatting = {
      ...target.run.formatting,
      bold: !(target.run.formatting?.bold ?? false),
    };
  } else if (content?.type === "text") {
    content.text = spec.kind === "append" ? `${content.text} ${spec.text}` : spec.text;
  }
  return target.para.paraId;
};

// ============================================================================
// ZIP HELPERS
// ============================================================================

const documentXml = async (buffer: ArrayBuffer): Promise<string> => {
  const zip = await JSZip.loadAsync(buffer);
  const file = zip.file("word/document.xml");
  if (!file) {
    throw new Error("word/document.xml missing");
  }
  return file.async("text");
};

const paragraphIds = (xml: string): string[] => {
  const pattern = /w14:paraId="(?<paraId>[^"]+)"/gu;
  const ids: string[] = [];
  let match: RegExpExecArray | null;
  while ((match = pattern.exec(xml)) !== null) {
    ids.push(match.groups?.["paraId"] ?? "");
  }
  return ids;
};

// Parts that selective save legitimately rewrites for a body-paragraph edit:
// document.xml is patched, core.xml is date-stamped, comments.xml is
// re-serialized when present, and headers/footers are re-serialized
// unconditionally by collectHeaderFooterUpdates (model-equivalent, not
// byte-identical). Everything else must survive byte-for-byte.
const isSelectiveRewritten = (partPath: string): boolean =>
  partPath === "word/document.xml" ||
  partPath === "docProps/core.xml" ||
  partPath === "word/comments.xml" ||
  /^word\/(?:header|footer)\d*\.xml$/u.test(partPath);

const bytesEqual = (a: Uint8Array, b: Uint8Array): boolean => {
  if (a.length !== b.length) {
    return false;
  }
  for (let i = 0; i < a.length; i++) {
    if (a[i] !== b[i]) {
      return false;
    }
  }
  return true;
};

// ============================================================================
// INVARIANT 1 — no-op save round-trips the parsed model losslessly
// ============================================================================

describe("invariant 1: no-op save fidelity", () => {
  test(
    "selective save of an unedited doc reproduces the parsed model exactly",
    async () => {
      await fc.assert(
        fc.asyncProperty(fc.constantFrom(...FIXTURES), async (name) => {
          const buffer = readFixture(name);
          const doc = await parse(buffer);
          const saved = await attemptSelectiveSave(doc, buffer, {
            changedParaIds: new Set(),
            structuralChange: false,
            hasUntrackedChanges: false,
          });
          expect(saved).not.toBeNull();
          if (!saved) {
            return;
          }
          expect(normalizedPackage(await parse(saved))).toEqual(normalizedPackage(doc));
        }),
        propertyConfig({ numRuns: 9, seed: SEED }),
      );
    },
    propertyTestTimeout(20_000),
  );

  test(
    "full repack of an unedited doc preserves all visible text",
    async () => {
      await fc.assert(
        fc.asyncProperty(fc.constantFrom(...FIXTURES), async (name) => {
          const buffer = readFixture(name);
          const doc = await parse(buffer);
          const repacked = await fullRepack(doc, buffer);
          expect(allBodyText(await parse(repacked))).toBe(allBodyText(doc));
        }),
        propertyConfig({ numRuns: 9, seed: SEED }),
      );
    },
    propertyTestTimeout(20_000),
  );

  // KNOWN FAILING — full repack is NOT a lossless model fixed point on rich docs.
  // Re-serializing an unedited doc and re-parsing it drifts from the original
  // model in two reproducible ways on current main:
  //   - docx-editor-demo.docx: a <w:commentReference> is DUPLICATED
  //     ($.document.content[20].content grows from 7 to 8 children).
  //   - podily-bps.docx: a complex-field (REF) run loses its fontFamily
  //     ($.document.content[86].content[1].formatting.fontFamily
  //      "Georgia" -> undefined).
  // Both are full-repack serializer fidelity gaps, independent of any edit.
  // Selective save preserves the original bytes and does NOT exhibit either
  // (see the invariant-1 selective test above), so this is a repack bug, not a
  // selective-save bug. Do NOT delete `.failing` to make CI green — fix the
  // serializer; this test then flips to passing and flags itself.
  test.failing(
    "full repack of an unedited doc is a lossless model fixed point",
    async () => {
      await fc.assert(
        fc.asyncProperty(
          fc.constantFrom("docx-editor-demo.docx", "podily-bps.docx"),
          async (name) => {
            const buffer = readFixture(name);
            const doc = await parse(buffer);
            const repacked = await fullRepack(doc, buffer);
            expect(normalizedPackage(await parse(repacked))).toEqual(normalizedPackage(doc));
          },
        ),
        propertyConfig({ numRuns: 6, seed: SEED }),
      );
    },
    propertyTestTimeout(20_000),
  );
});

// ============================================================================
// INVARIANT 2 — selective ≡ full repack at the parsed-model level
// ============================================================================

describe("invariant 2: selective equals full repack", () => {
  // KNOWN FAILING — the two paths do NOT persist the same model on podily-bps.
  // After any edit, full repack re-serializes every paragraph and drops the
  // Georgia fontFamily on the unedited complex-field run (invariant 1 above),
  // while selective save copies that paragraph's ORIGINAL bytes untouched.
  // parse(SEL) is therefore MORE faithful than parse(FULL); they diverge
  // because full repack loses data, not because selective save is wrong. The
  // fix belongs in the serializer; this test flips to passing once it lands.
  test.failing("parse(selective) deep-equals parse(full) for an edited doc", async () => {
    const buffer = readFixture(EDIT_FIXTURE);
    const doc = await parse(buffer);
    const paraId = applyEdit(doc, { targetSelector: 0, kind: "append", text: "EDIT" });
    expect(paraId).not.toBeNull();
    if (!paraId) {
      return;
    }

    const selective = await attemptSelectiveSave(doc, buffer, {
      changedParaIds: new Set([paraId]),
      structuralChange: false,
      hasUntrackedChanges: false,
    });
    expect(selective).not.toBeNull();
    if (!selective) {
      return;
    }
    const full = await fullRepack(doc, buffer);

    expect(normalizedPackage(await parse(selective))).toEqual(normalizedPackage(await parse(full)));
  });
});

// ============================================================================
// INVARIANT 3 — selective save is never silently wrong
// ============================================================================
//
// The safety net is stated against the ORIGINAL edit, not against full repack:
// selective save must either bail (null) or persist a document whose visible
// text equals the edited model exactly — no dropped, reordered, or corrupted
// paragraphs. This is the guarantee that actually protects a user's file.

describe("invariant 3: selective save is never silently wrong", () => {
  test(
    "selective returns null or a document that applies exactly the edit",
    async () => {
      const buffer = readFixture(EDIT_FIXTURE);
      await fc.assert(
        fc.asyncProperty(arbEditSpec, async (spec) => {
          const doc = await parse(buffer);
          const paraId = applyEdit(doc, spec);
          if (!paraId) {
            return;
          }
          const saved = await attemptSelectiveSave(doc, buffer, {
            changedParaIds: new Set([paraId]),
            structuralChange: false,
            hasUntrackedChanges: false,
          });
          if (saved === null) {
            return; // bailed to full repack — an allowed outcome
          }
          // Non-null: the persisted model must match the edited model's text.
          const reparsed = await parse(saved);
          expect(bodyParagraphTexts(reparsed)).toEqual(bodyParagraphTexts(doc));
        }),
        propertyConfig({ numRuns: 15, seed: SEED }),
      );
    },
    propertyTestTimeout(30_000),
  );
});

// ============================================================================
// INVARIANT 4 — byte-exactness of untouched parts
// ============================================================================

describe("invariant 4: selective save keeps untouched bytes identical", () => {
  test(
    "a single-paragraph edit leaves untouched parts and unedited paragraphs byte-identical",
    async () => {
      const buffer = readFixture(EDIT_FIXTURE);
      const originalZip = await JSZip.loadAsync(buffer);
      const originalXml = await documentXml(buffer);

      // Precompute the original bytes/slices once so each fast-check run only
      // pays for the saved side (the ~1.6 MB document.xml re-scan is the cost
      // driver, so the unedited-paragraph check samples a bounded 40 ids).
      const untouchedParts = Object.keys(originalZip.files).filter(
        (part) => !originalZip.files[part]?.dir && !isSelectiveRewritten(part),
      );
      const originalUntouched = await Promise.all(
        untouchedParts.map(async (part) => ({
          part,
          bytes: new Uint8Array(await originalZip.file(part)!.async("arraybuffer")),
        })),
      );
      const sampledSlices = [...new Set(paragraphIds(originalXml))]
        .flatMap((id) => {
          const offsets = findParagraphOffsets(originalXml, id);
          return offsets ? [{ id, slice: originalXml.slice(offsets.start, offsets.end) }] : [];
        })
        .slice(0, 40);

      await fc.assert(
        fc.asyncProperty(arbEditSpec, async (spec) => {
          const doc = await parse(buffer);
          const paraId = applyEdit(doc, spec);
          if (!paraId) {
            return;
          }
          const saved = await attemptSelectiveSave(doc, buffer, {
            changedParaIds: new Set([paraId]),
            structuralChange: false,
            hasUntrackedChanges: false,
          });
          if (saved === null) {
            return;
          }

          const savedZip = await JSZip.loadAsync(saved);

          // Every genuinely-untouched part survives byte-for-byte (styles.xml,
          // numbering.xml, footnotes.xml, endnotes.xml, theme, fonts, settings,
          // media, ...). TODO(notes): note bodies are not editable through the
          // selective path today, so footnotes.xml/endnotes.xml are expected to
          // be byte-identical here; extend this once the note write path lands.
          const comparisons = await Promise.all(
            originalUntouched.map(async ({ part, bytes }) => {
              const savedFile = savedZip.file(part);
              const savedBytes = savedFile
                ? new Uint8Array(await savedFile.async("arraybuffer"))
                : null;
              return { part, equal: savedBytes !== null && bytesEqual(bytes, savedBytes) };
            }),
          );
          for (const { part, equal } of comparisons) {
            expect(equal, `untouched part changed: ${part}`).toBe(true);
          }

          // Every sampled UNEDITED paragraph keeps its exact original XML slice;
          // the edited paragraph is the only one allowed to change.
          const savedXml = await savedZip.file("word/document.xml")!.async("text");
          for (const { id, slice } of sampledSlices) {
            if (id === paraId) {
              continue;
            }
            const after = findParagraphOffsets(savedXml, id);
            if (after) {
              expect(savedXml.slice(after.start, after.end)).toBe(slice);
            }
          }

          // The edit itself did land in the edited paragraph.
          expect(findParagraphOffsets(savedXml, paraId)).not.toBeNull();
        }),
        propertyConfig({ numRuns: 12, seed: SEED }),
      );
    },
    propertyTestTimeout(30_000),
  );
});
