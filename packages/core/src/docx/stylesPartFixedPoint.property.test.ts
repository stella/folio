/**
 * Saving must not grow `word/styles.xml`.
 *
 * The part is preserved verbatim and the styles the model added since are
 * appended before its root close, so "which ids does the part already define"
 * decides whether a style is written again. That question was answered by
 * scanning the part's text for `w:styleId="…"`, and the text is the XML
 * spelling of an id while the model holds its decoded value: a single-quoted
 * attribute, an escaped character, or an earlier attribute containing `>` all
 * read as a style the part lacks. It was then appended on every save, so save
 * n carried n copies.
 *
 * The property generates style tables rather than one fixture because the
 * variable is the spelling, not the style.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";

import { escapeXmlAttribute } from "@stll/docx-core";

import { propertyConfig } from "../../../../test/property-testing";

import { parseDocx } from "./parser";
import { createEmptyDocx, repackDocx } from "./rezip";
import type { Document, Style } from "../types/document";
import {
  findChildByNamespaceUri,
  findChildrenByNamespaceUri,
  getAttributeByNamespaceUri,
  parseXml,
  WORDPROCESSINGML_NAMESPACE_URIS,
} from "./xmlParser";

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const W_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/**
 * Ids a real style table carries, every one of them legal `ST_String`.
 *
 * `Header & Footer` is the one the public corpus found, on a Pages export.
 */
const STYLE_IDS = [
  "Normal",
  "Body A",
  "Header & Footer",
  "Quote<Strong>",
  'Say "Hi"',
  "Nadpis-Ú",
] as const;

type StyleSpelling = {
  id: string;
  /** Attribute delimiter. Both are legal XML; only one was ever read. */
  quote: '"' | "'";
  /** An earlier attribute whose value carries a raw `>`, also legal XML. */
  aliasWithAngle: boolean;
};

const styleElement = ({ id, quote, aliasWithAngle }: StyleSpelling): string => {
  // A raw `>` is legal inside an attribute value, and the escaper would take
  // it away, so this one is written as the producer wrote it.
  const alias = aliasWithAngle ? ` w:aliases="before>after"` : "";
  const escaped = escapeXmlAttribute(id);
  return (
    `<w:style w:type="paragraph"${alias} w:styleId=${quote}${escaped}${quote}>` +
    `<w:name w:val=${quote}${escaped}${quote}/></w:style>`
  );
};

const stylesPart = (spellings: readonly StyleSpelling[]): string =>
  `${XML_DECLARATION}<w:styles xmlns:w="${W_NAMESPACE}">` +
  `${spellings.map(styleElement).join("")}</w:styles>`;

const buildDocx = async (spellings: readonly StyleSpelling[]): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(await createEmptyDocx());
  zip.file("word/styles.xml", stylesPart(spellings));
  return zip.generateAsync({ type: "arraybuffer" });
};

/** The ids a written `word/styles.xml` defines, in order, duplicates kept. */
const writtenStyleIds = async (buffer: ArrayBuffer): Promise<string[]> => {
  const xml = await (await JSZip.loadAsync(buffer)).file("word/styles.xml")?.async("text");
  if (xml === undefined) {
    throw new Error("the saved package has no styles part");
  }
  const root = parseXml(xml);
  const styles = findChildByNamespaceUri(root, WORDPROCESSINGML_NAMESPACE_URIS, "styles") ?? root;
  return findChildrenByNamespaceUri(styles, WORDPROCESSINGML_NAMESPACE_URIS, "style").map(
    (style) => getAttributeByNamespaceUri(style, WORDPROCESSINGML_NAMESPACE_URIS, "styleId") ?? "",
  );
};

const partText = async (buffer: ArrayBuffer, path: string): Promise<string | undefined> =>
  (await JSZip.loadAsync(buffer)).file(path)?.async("text");

/** Every part of a package, by path, so a byte comparison names the part. */
const packageParts = async (buffer: ArrayBuffer): Promise<Map<string, string>> => {
  const zip = await JSZip.loadAsync(buffer);
  return new Map(
    await Promise.all(
      Object.entries(zip.files)
        .filter(([, file]) => !file.dir)
        .map(
          async ([path, file]): Promise<readonly [string, string]> => [
            path,
            await file.async("base64"),
          ],
        ),
    ),
  );
};

const addStyles = (document: Document, ids: readonly string[]): void => {
  const styles = document.package.styles;
  if (!styles) {
    throw new Error("the parsed package has no style table");
  }
  styles.styles = [
    ...styles.styles,
    ...ids.map((styleId): Style => ({ styleId, type: "paragraph", name: styleId })),
  ];
};

const styleTable = fc.uniqueArray(
  fc.record({
    id: fc.constantFrom(...STYLE_IDS),
    quote: fc.constantFrom<'"' | "'">('"', "'"),
    aliasWithAngle: fc.boolean(),
  }),
  { minLength: 1, maxLength: STYLE_IDS.length, selector: ({ id }) => id },
);

/**
 * What the editor added since: ids drawn from the same pool, so a generated
 * case may add a style the part already defines (which must not be written
 * again) or the same style twice (which must not be written twice).
 */
const editorAdditions = fc.array(fc.constantFrom(...STYLE_IDS), { maxLength: 4 });

describe("the styles part is a fixed point across saves", () => {
  test("no style id is written twice, and the second save changes nothing", async () => {
    await fc.assert(
      fc.asyncProperty(styleTable, editorAdditions, async (spellings, additions) => {
        const parsed = await parseDocx(await buildDocx(spellings), { preloadFonts: false });
        addStyles(parsed, additions);

        const first = await repackDocx(parsed, { updateModifiedDate: false });
        const reparsed = await parseDocx(first, { preloadFonts: false });
        const second = await repackDocx(reparsed, { updateModifiedDate: false });

        const ids = await writtenStyleIds(first);
        expect(ids).toEqual([...new Set(ids)]);
        // Every id the model holds is defined, so convergence is not achieved
        // by dropping a style.
        for (const { id } of spellings) {
          expect(ids).toContain(id);
        }
        for (const id of additions) {
          expect(ids).toContain(id);
        }

        const firstParts = await packageParts(first);
        const secondParts = await packageParts(second);
        expect([...secondParts.keys()].sort()).toEqual([...firstParts.keys()].sort());
        for (const [path, content] of firstParts) {
          expect({ path, content }).toEqual({ path, content: secondParts.get(path) });
        }
      }),
      propertyConfig({ numRuns: 40 }),
    );
  });

  test("repacking an unedited document leaves the styles part byte-identical", async () => {
    // The construct the corpus minimised to: an id whose XML spelling is not
    // its value. Reading the spelling reported the style as missing.
    const spellings: StyleSpelling[] = [
      { id: "Normal", quote: '"', aliasWithAngle: false },
      { id: "Header & Footer", quote: '"', aliasWithAngle: false },
    ];
    const source = await buildDocx(spellings);
    const parsed = await parseDocx(source, { preloadFonts: false });

    const saved = await repackDocx(parsed, { updateModifiedDate: false });

    expect(await partText(saved, "word/styles.xml")).toBe(stylesPart(spellings));
  });
});
