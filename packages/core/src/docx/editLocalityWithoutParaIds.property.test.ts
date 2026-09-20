/**
 * One edit stays one edit, whether or not the producer wrote paragraph ids.
 *
 * Word 2010+ stamps a `w14:paraId` on every paragraph and the selective save
 * keys paragraph identity on it. LibreOffice, Google Docs, python-docx and
 * docx4j write none, so folio mints one for every paragraph at parse — and a
 * minted id names nothing in the file. Keyed on ids alone the save then finds
 * *every* paragraph unaddressable, declines the splice, and rebuilds the whole
 * part, so a one-character edit rewrites the document.
 *
 * These properties fix the three id regimes a real package can be in — none,
 * all, and the mixed one Word produces when it rewrites part of a foreign
 * document — and assert what the reviewer owes each of them:
 *
 *   1. A text replacement leaves every other paragraph's bytes alone, and
 *      leaves the part's id coverage exactly as the author wrote it.
 *   2. Every edit kind, structural ones included, saves a package that parses,
 *      declares the namespace of every id it writes, and still holds the
 *      paragraphs it never named.
 *
 * Property 1 is the guard for the corpus `edit-locality` family's largest
 * signature (`content[].paraId: absent became "<hex>"`). A structural edit
 * takes the full-repack path by construction — the change tracker routes it
 * there before any splice is attempted — so property 1 is stated for the
 * replacement that does reach the selective path, and property 2 covers the
 * rest.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import { FolioDocxReviewer } from "../ai-edits/headless";
import type { FolioAIBlock, FolioAIEditOperation } from "../ai-edits/types";
import { parseDocx } from "./parser";
import { parseXmlDocument } from "./xmlParser";

setDefaultTimeout(propertyTestTimeout(30_000));

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`;

const ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`;

const DOCUMENT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml";

/** How many of the source's paragraphs carry an authored `w14:paraId`. */
const ID_COVERAGE = { none: "none", all: "all", mixed: "mixed" } as const;
type IdCoverage = (typeof ID_COVERAGE)[keyof typeof ID_COVERAGE];

const EDIT_KIND = {
  replace: "replace",
  split: "split",
  merge: "merge",
  insert: "insert",
  delete: "delete",
} as const;
type EditKind = (typeof EDIT_KIND)[keyof typeof EDIT_KIND];

/** A character no generated body text contains, so it names the edited text. */
const MARKER = "‸";

/** Every run zips, opens, edits and saves two packages; the default 5s is not it. */
const PROPERTY_TIMEOUT_MS = 180_000;

const paragraphTextArb = fc.stringMatching(/^[A-Za-z]{2,10} [A-Za-z]{2,10}$/u);

const documentArb = fc.record({
  texts: fc.uniqueArray(paragraphTextArb, { minLength: 3, maxLength: 6 }),
  coverage: fc.constantFrom(...Object.values(ID_COVERAGE)),
});

/** Word writes ids on what it rewrote, so mixed coverage is a real regime. */
const carriesId = (coverage: IdCoverage, ordinal: number): boolean => {
  switch (coverage) {
    case ID_COVERAGE.none:
      return false;
    case ID_COVERAGE.all:
      return true;
    case ID_COVERAGE.mixed:
      return ordinal % 2 === 0;
    default: {
      const unreachable: never = coverage;
      return unreachable;
    }
  }
};

const authoredId = (ordinal: number): string =>
  (0x11110000 + ordinal * 0x1111).toString(16).toUpperCase().padStart(8, "0");

const buildDocumentXml = (texts: readonly string[], coverage: IdCoverage): string => {
  const declaresW14 = coverage !== ID_COVERAGE.none;
  const body = texts
    .map((text, ordinal) => {
      const id = carriesId(coverage, ordinal)
        ? ` w14:paraId="${authoredId(ordinal)}" w14:textId="${authoredId(ordinal)}"`
        : "";
      return `<w:p${id}><w:r><w:t>${text}</w:t></w:r></w:p>`;
    })
    .join("");
  const w14 = declaresW14 ? ` xmlns:w14="${W14_NS}"` : "";
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document xmlns:w="${W_NS}"${w14}><w:body>${body}<w:sectPr/></w:body></w:document>`;
};

const buildPackage = async (
  texts: readonly string[],
  coverage: IdCoverage,
): Promise<ArrayBuffer> => {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", CONTENT_TYPES);
  zip.file("_rels/.rels", ROOT_RELS);
  zip.file("word/_rels/document.xml.rels", DOCUMENT_RELS);
  zip.file("word/document.xml", buildDocumentXml(texts, coverage));
  const bytes = await zip.generateAsync({ type: "uint8array" });
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
};

const documentPartOf = async (docx: ArrayBuffer): Promise<string> => {
  const entry = (await JSZip.loadAsync(docx)).file("word/document.xml");
  if (!entry) {
    throw new Error("saved package has no word/document.xml");
  }
  return entry.async("text");
};

/**
 * The `<w:p>…</w:p>` slices of a part, in document order.
 *
 * Deliberately not the patcher's own scanner: a test that measured locality
 * with the scanner under test would agree with it by construction. The
 * generated bodies are flat, so a forward split on the paragraph tags is the
 * whole of what this needs.
 */
const paragraphBytes = (xml: string): string[] =>
  [...xml.matchAll(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>|<w:p(?:\s[^>]*)?\/>/gu)].map(
    (match) => match[0],
  );

const paragraphTexts = (xml: string): string[] =>
  paragraphBytes(xml).map((paragraph) =>
    [...paragraph.matchAll(/<w:t(?:\s[^>]*)?>(?<text>[^<]*)<\/w:t>/gu)]
      .map((match) => match.groups?.["text"] ?? "")
      .join(""),
  );

const paraIdCount = (xml: string): number => [...xml.matchAll(/\sw14:paraId="/gu)].length;

/** Every id the part writes must resolve, or the part is not a valid package. */
const declaresEveryIdNamespace = (xml: string): boolean =>
  paraIdCount(xml) === 0 || xml.includes(`xmlns:w14="${W14_NS}"`);

const isSubsequence = (needles: readonly string[], haystack: readonly string[]): boolean => {
  let cursor = 0;
  for (const needle of needles) {
    cursor = haystack.indexOf(needle, cursor) + 1;
    if (cursor === 0) {
      return false;
    }
  }
  return true;
};

type EditPlan = { operation: FolioAIEditOperation; untouched: string[] };

/** The operation for `kind`, and the texts it must leave standing. */
const planEdit = (kind: EditKind, blocks: readonly FolioAIBlock[]): EditPlan => {
  // SAFETY: the generator produces at least three paragraphs.
  const target = blocks[1]!;
  const others = blocks.filter((block) => block.id !== target.id).map((block) => block.text);
  switch (kind) {
    case EDIT_KIND.replace:
      return {
        operation: {
          id: "edit",
          type: "replaceInBlock",
          blockId: target.id,
          find: target.text,
          replace: `${target.text}${MARKER}`,
        },
        untouched: others,
      };
    case EDIT_KIND.split:
      return {
        operation: {
          id: "edit",
          type: "splitBlock",
          blockId: target.id,
          offset: Math.max(1, Math.floor(target.text.length / 2)),
        },
        untouched: others,
      };
    case EDIT_KIND.merge:
      return {
        operation: { id: "edit", type: "mergeBlockWithNext", blockId: target.id },
        untouched: blocks
          .filter((block, index) => index !== 1 && index !== 2)
          .map((block) => block.text),
      };
    case EDIT_KIND.insert:
      return {
        operation: {
          id: "edit",
          type: "insertAfterBlock",
          blockId: target.id,
          text: `inserted${MARKER}`,
        },
        untouched: blocks.map((block) => block.text),
      };
    case EDIT_KIND.delete:
      return {
        operation: { id: "edit", type: "deleteBlock", blockId: target.id },
        untouched: others,
      };
    default: {
      const unreachable: never = kind;
      return unreachable;
    }
  }
};

type SavedEdit = { buffer: ArrayBuffer; control: string; edited: string; untouched: string[] };

const saveEdited = async (
  texts: readonly string[],
  coverage: IdCoverage,
  kind: EditKind,
): Promise<SavedEdit> => {
  const source = await buildPackage(texts, coverage);
  const reviewer = await FolioDocxReviewer.fromBuffer(source);
  const { operation, untouched } = planEdit(kind, reviewer.snapshot().blocks);
  const applied = reviewer.applyOperations([operation], { mode: "direct" });
  expect(applied.applied.length).toBe(1);
  const buffer = await reviewer.toBuffer();
  // The control save pays the same package-wide id and version normalization
  // the edited save does, so what survives the comparison is the edit's reach.
  const control = await documentPartOf(
    await (await FolioDocxReviewer.fromBuffer(source)).toBuffer(),
  );
  return { buffer, control, edited: await documentPartOf(buffer), untouched };
};

describe("an edit to one paragraph stays local without authored paraIds", () => {
  test(
    "a text replacement leaves every other paragraph byte-exact",
    async () => {
      await fc.assert(
        fc.asyncProperty(documentArb, async ({ texts, coverage }) => {
          const { control, edited } = await saveEdited(texts, coverage, EDIT_KIND.replace);

          const before = paragraphBytes(control);
          const after = paragraphBytes(edited);
          expect(after.length).toBe(before.length);
          for (const [ordinal, paragraph] of before.entries()) {
            if (paragraph.includes(MARKER) || after[ordinal]?.includes(MARKER)) {
              continue;
            }
            expect(after[ordinal]).toBe(paragraph);
          }
          // The edit landed, and it landed in exactly one paragraph.
          expect(after.filter((paragraph) => paragraph.includes(MARKER))).toHaveLength(1);
          // An id the author never wrote is a model address, not a package fact:
          // saving must not promote it into the file.
          expect(paraIdCount(edited)).toBe(paraIdCount(control));
        }),
        propertyConfig({ numRuns: 60 }),
      );
    },
    PROPERTY_TIMEOUT_MS,
  );

  test(
    "every edit kind saves a valid package that keeps its untouched paragraphs",
    async () => {
      await fc.assert(
        fc.asyncProperty(
          documentArb,
          fc.constantFrom(...Object.values(EDIT_KIND)),
          async ({ texts, coverage }, kind) => {
            const { buffer, edited, untouched } = await saveEdited(texts, coverage, kind);

            expect(parseXmlDocument(edited)).not.toBeNull();
            expect(declaresEveryIdNamespace(edited)).toBe(true);
            expect(isSubsequence(untouched, paragraphTexts(edited))).toBe(true);

            const reparsed = await parseDocx(buffer, { preloadFonts: false });
            expect(
              reparsed.package.document.content.filter((block) => block.type === "paragraph"),
            ).not.toHaveLength(0);
          },
        ),
        propertyConfig({ numRuns: 60 }),
      );
    },
    PROPERTY_TIMEOUT_MS,
  );
});
