/**
 * A source-XML visibility oracle for nested field-code structures.
 *
 * `w:fldChar` explicitly supports nesting one field inside another, and only
 * the region between a field's own `separate` and `end` is displayed — its
 * `begin`..`separate` instruction region never is, whatever it contains,
 * including another field's entire begin/separate/result/end. A run of text
 * is visible only if it sits in the result region of *every* field it is
 * nested inside; one enclosing instruction region is enough to hide it, no
 * matter how many levels deep or which of those levels itself has already
 * passed its own separator.
 *
 * `ai-edits/field-results.test.ts` and `ai-edits/story-text-parity.test.ts`
 * cover only non-nested fields — a single field with a cached result, never
 * one field inside another. This generates a random chain of nested fields
 * (depth 1-3), each link placed in its parent's still-open instruction region
 * or in its already-open result region, with a unique text sentinel standing
 * in for that link's own cached result, and checks every reader that walks
 * parsed field content — `getContent` (the ProseMirror-loaded projection),
 * `listStories`'s unloaded model walk (a footnote, read before it is ever
 * loaded into an editor state), and `docxToMarkdown` (a fully independent
 * parse) — shows exactly the sentinels whose whole ancestor chain has passed
 * its own separator, and none of the rest.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import { RELATIONSHIP_TYPES } from "./relsParser";
import { docxToMarkdown } from "../server";
import { FolioDocxReviewer } from "../ai-edits/headless";

setDefaultTimeout(propertyTestTimeout(30_000));

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const MAX_DEPTH = 3;

type LinkPosition = "code" | "result";
type LeafKind = "complex" | "simple";

type GeneratedCase = {
  depth: number;
  /** `positions[k]` is where link `k + 1` sits inside link `k`. */
  positions: LinkPosition[];
  leafKind: LeafKind;
  /** Split the deepest link's own instruction/result text across two runs. */
  splitLeafRuns: boolean;
};

const generatedCase: fc.Arbitrary<GeneratedCase> = fc
  .integer({ min: 1, max: MAX_DEPTH })
  .chain((depth) =>
    fc.record({
      depth: fc.constant(depth),
      positions: fc.array(fc.constantFrom<LinkPosition>("code", "result"), {
        minLength: depth - 1,
        maxLength: depth - 1,
      }),
      leafKind: fc.constantFrom<LeafKind>("complex", "simple"),
      splitLeafRuns: fc.boolean(),
    }),
  );

const resultSentinel = (level: number): string => `RESULT_${level}_TOKEN`;
const codeMarker = (level: number): string => `CODE_${level}_MARKER`;

const run = (text: string): string => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const instrRun = (text: string): string =>
  `<w:r><w:instrText xml:space="preserve">${text}</w:instrText></w:r>`;
const fldChar = (charType: "begin" | "separate" | "end"): string =>
  `<w:r><w:fldChar w:fldCharType="${charType}"/></w:r>`;
const complexField = (code: string, result: string): string =>
  fldChar("begin") + code + fldChar("separate") + result + fldChar("end");

/**
 * Build the outermost link's XML for a generated case, and say which levels'
 * result sentinels a reader must show.
 */
const buildNestedFieldXml = ({
  depth,
  positions,
  leafKind,
  splitLeafRuns,
}: GeneratedCase): { xml: string; visibleLevels: boolean[] } => {
  const visibleLevels: boolean[] = [];
  let ancestorsAllPastSeparator = true;
  for (let level = 0; level < depth; level++) {
    visibleLevels.push(ancestorsAllPastSeparator);
    if (level < depth - 1) {
      ancestorsAllPastSeparator &&= positions[level] === "result";
    }
  }

  let childXml = "";
  let childPosition: LinkPosition | null = null;
  for (let level = depth - 1; level >= 0; level--) {
    const isLeaf = level === depth - 1;
    const sentinel = resultSentinel(level);
    const resultText =
      isLeaf && splitLeafRuns ? run(sentinel.slice(0, 3)) + run(sentinel.slice(3)) : run(sentinel);
    let levelXml: string;
    if (isLeaf && leafKind === "simple") {
      levelXml = `<w:fldSimple w:instr=" DOCPROPERTY X${level} ">${resultText}</w:fldSimple>`;
    } else {
      const instrText =
        isLeaf && splitLeafRuns
          ? instrRun(` ${codeMarker(level)} `.slice(0, 8)) +
            instrRun(` ${codeMarker(level)} `.slice(8))
          : instrRun(` ${codeMarker(level)} `);
      const before = childPosition === "code" ? childXml : "";
      const after = childPosition === "result" ? childXml : "";
      levelXml = complexField(instrText + before, resultText + after);
    }
    childXml = levelXml;
    childPosition = level > 0 ? positions[level - 1] : null;
  }

  return { xml: childXml, visibleLevels };
};

const createDocx = async (fieldXml: string): Promise<ArrayBuffer> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `${XML_DECLARATION}
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
</Types>`,
  );
  zip.file(
    "_rels/.rels",
    `${XML_DECLARATION}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${RELATIONSHIP_TYPES.officeDocument}" Target="word/document.xml"/>
</Relationships>`,
  );
  zip.file(
    "word/_rels/document.xml.rels",
    `${XML_DECLARATION}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId30" Type="${RELATIONSHIP_TYPES.footnotes}" Target="footnotes.xml"/>
</Relationships>`,
  );
  const fieldParagraph = `<w:p>${run("Before ")}${fieldXml}${run(" after.")}</w:p>`;
  zip.file(
    "word/footnotes.xml",
    `${XML_DECLARATION}
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:id="2">${fieldParagraph}</w:footnote>
</w:footnotes>`,
  );
  zip.file(
    "word/document.xml",
    `${XML_DECLARATION}
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    ${fieldParagraph}
    <w:p><w:r><w:footnoteReference w:id="2"/></w:r><w:r><w:t>anchor</w:t></w:r></w:p>
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
  </w:body>
</w:document>`,
  );
  zip.file(
    "word/styles.xml",
    `${XML_DECLARATION}
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`,
  );
  return zip.generateAsync({ type: "arraybuffer" });
};

/** Every sentinel a text blob contains, so "none of the rest" is one diff. */
const sentinelsPresent = (text: string, depth: number): number[] =>
  Array.from({ length: depth }, (_, level) => level).filter((level) =>
    text.includes(resultSentinel(level)),
  );

describe("nested field-code visibility oracle", () => {
  test("every reader shows exactly the result-region sentinels", async () => {
    await fc.assert(
      fc.asyncProperty(generatedCase, async (generated) => {
        const { xml, visibleLevels } = buildNestedFieldXml(generated);
        const expected = visibleLevels
          .map((visible, level) => (visible ? level : null))
          .filter((level): level is number => level !== null);

        const buffer = await createDocx(xml);

        const reviewer = await FolioDocxReviewer.fromBuffer(buffer);
        const mainText = reviewer.getContent().at(0)?.text ?? "";
        expect(sentinelsPresent(mainText, generated.depth)).toEqual(expected);

        const footnote = reviewer.listStories().find(({ handle }) => handle.type === "footnote");
        expect(sentinelsPresent(footnote?.text ?? "", generated.depth)).toEqual(expected);

        const markdown = await docxToMarkdown(buffer);
        expect(sentinelsPresent(markdown, generated.depth)).toEqual(expected);
      }),
      propertyConfig({ numRuns: 200 }),
    );
  }, 180_000);
});
