import { describe, expect, test } from "bun:test";

import {
  assessFontEnvironment,
  attributeDivergences,
  clusterCorpus,
  computeFontSubstitutionTags,
  fontFamiliesMatch,
  scanDocumentXml,
} from "../features";
import type { DocFeatures, ParagraphFeatures } from "../features";
import type { Divergence, DocGeom, FeatureAttributedResult, ParityResult } from "../types";

const wrapBody = (body: string): string =>
  `<?xml version="1.0"?><w:document><w:body>${body}</w:body></w:document>`;

describe("scanDocumentXml: paragraph scanning + feature tagging", () => {
  test("plain paragraph carries no feature tags", () => {
    const { paragraphs } = scanDocumentXml(
      wrapBody("<w:p><w:r><w:t>Hello world</w:t></w:r></w:p>"),
    );
    expect(paragraphs).toEqual([{ normText: "Hello world", features: [] }]);
  });

  test("paragraph inside a table is tagged 'table', not 'nested-table'", () => {
    const xml = wrapBody(
      "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>Cell text</w:t></w:r></w:p></w:tc></w:tr></w:tbl>",
    );
    const { paragraphs } = scanDocumentXml(xml);
    expect(paragraphs).toHaveLength(1);
    expect(paragraphs[0]?.normText).toBe("Cell text");
    expect(paragraphs[0]?.features).toEqual(["table"]);
  });

  test("paragraph inside a nested table is tagged both 'table' and 'nested-table'", () => {
    const xml = wrapBody(
      "<w:tbl><w:tr><w:tc>" +
        "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>Nested cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>" +
        "</w:tc></w:tr></w:tbl>",
    );
    const { paragraphs } = scanDocumentXml(xml);
    expect(paragraphs).toHaveLength(1);
    expect(paragraphs[0]?.normText).toBe("Nested cell");
    expect(paragraphs[0]?.features).toEqual(["table", "nested-table"]);
  });

  test("paragraph after a closed table is not tagged 'table'", () => {
    const xml = wrapBody(
      "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>In table</w:t></w:r></w:p></w:tc></w:tr></w:tbl>" +
        "<w:p><w:r><w:t>After table</w:t></w:r></w:p>",
    );
    const { paragraphs } = scanDocumentXml(xml);
    expect(paragraphs[0]?.features).toEqual(["table"]);
    expect(paragraphs[1]?.features).toEqual([]);
  });

  test("a run's literal tab char is tagged 'tab', a tab-stop def is tagged 'tab-stops', independently", () => {
    const tabChar = scanDocumentXml(
      wrapBody("<w:p><w:r><w:t>a</w:t><w:tab/><w:t>b</w:t></w:r></w:p>"),
    ).paragraphs[0];
    expect(tabChar?.features).toEqual(["tab"]);

    const tabStops = scanDocumentXml(
      wrapBody(
        '<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="720"/></w:tabs></w:pPr><w:r><w:t>x</w:t></w:r></w:p>',
      ),
    ).paragraphs[0];
    expect(tabStops?.features).toEqual(["tab-stops"]);

    const both = scanDocumentXml(
      wrapBody(
        '<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="720"/></w:tabs></w:pPr>' +
          "<w:r><w:t>a</w:t><w:tab/><w:t>b</w:t></w:r></w:p>",
      ),
    ).paragraphs[0];
    expect(both?.features).toEqual(["tab-stops", "tab"]);
  });

  test("numbering (w:numPr) is tagged 'numbering'", () => {
    const xml = wrapBody(
      '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Item</w:t></w:r></w:p>',
    );
    expect(scanDocumentXml(xml).paragraphs[0]?.features).toEqual(["numbering"]);
  });

  test("justification variants map to justify / align-center / align-right", () => {
    const jc = (val: string) =>
      scanDocumentXml(
        wrapBody(`<w:p><w:pPr><w:jc w:val="${val}"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>`),
      ).paragraphs[0]?.features;
    expect(jc("both")).toEqual(["justify"]);
    expect(jc("distribute")).toEqual(["justify"]);
    expect(jc("center")).toEqual(["align-center"]);
    expect(jc("right")).toEqual(["align-right"]);
    expect(jc("end")).toEqual(["align-right"]);
  });

  test("spacing lineRule=atLeast/exact are tagged distinctly, auto+line is 'spacing-multiple'", () => {
    const atLeast = scanDocumentXml(
      wrapBody(
        '<w:p><w:pPr><w:spacing w:lineRule="atLeast" w:line="360"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>',
      ),
    ).paragraphs[0];
    expect(atLeast?.features).toEqual(["spacing-atLeast"]);

    const exact = scanDocumentXml(
      wrapBody(
        '<w:p><w:pPr><w:spacing w:lineRule="exact" w:line="240"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>',
      ),
    ).paragraphs[0];
    expect(exact?.features).toEqual(["spacing-exact"]);

    const multiple = scanDocumentXml(
      wrapBody(
        '<w:p><w:pPr><w:spacing w:lineRule="auto" w:line="480"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>',
      ),
    ).paragraphs[0];
    expect(multiple?.features).toEqual(["spacing-multiple"]);
  });

  test("anchored drawing is 'float-anchor', inline drawing is 'inline-image'", () => {
    const anchor = scanDocumentXml(
      wrapBody("<w:p><w:r><w:drawing><wp:anchor>stuff</wp:anchor></w:drawing></w:r></w:p>"),
    ).paragraphs[0];
    expect(anchor?.features).toEqual(["float-anchor"]);

    const inline = scanDocumentXml(
      wrapBody("<w:p><w:r><w:drawing><wp:inline>stuff</wp:inline></w:drawing></w:r></w:p>"),
    ).paragraphs[0];
    expect(inline?.features).toEqual(["inline-image"]);
  });

  test("field chars/instrText/fldSimple are tagged 'field'", () => {
    const fldChar = scanDocumentXml(
      wrapBody(
        '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>',
      ),
    ).paragraphs[0];
    expect(fldChar?.features).toEqual(["field"]);

    const fldSimple = scanDocumentXml(
      wrapBody('<w:p><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p>'),
    ).paragraphs[0];
    expect(fldSimple?.features).toEqual(["field"]);
  });

  test("hyperlink is tagged 'hyperlink' and its run text is captured", () => {
    const xml = wrapBody(
      '<w:p><w:hyperlink r:id="rId4"><w:r><w:t>Link text</w:t></w:r></w:hyperlink></w:p>',
    );
    const paragraph = scanDocumentXml(xml).paragraphs[0];
    expect(paragraph?.features).toEqual(["hyperlink"]);
    expect(paragraph?.normText).toBe("Link text");
  });

  test("page-break / column-break are distinguished by w:br w:type", () => {
    const pageBreak = scanDocumentXml(wrapBody('<w:p><w:r><w:br w:type="page"/></w:r></w:p>'))
      .paragraphs[0];
    expect(pageBreak?.features).toEqual(["page-break"]);

    const columnBreak = scanDocumentXml(wrapBody('<w:p><w:r><w:br w:type="column"/></w:r></w:p>'))
      .paragraphs[0];
    expect(columnBreak?.features).toEqual(["column-break"]);
  });

  test("a mid-document w:sectPr tags the containing paragraph 'sect-props' and the doc 'multi-section' at 2+", () => {
    const xml = wrapBody(
      '<w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr><w:r><w:t>Section break para</w:t></w:r></w:p>' +
        "<w:p><w:r><w:t>after</w:t></w:r></w:p>" +
        '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>',
    );
    const { paragraphs, bodyFeatures } = scanDocumentXml(xml);
    expect(paragraphs[0]?.features).toEqual(["sect-props"]);
    expect(paragraphs[1]?.features).toEqual([]);
    expect(bodyFeatures).toContain("multi-section");
  });

  test("a single body-level w:sectPr does not trigger 'multi-section'", () => {
    const xml = wrapBody(
      '<w:p><w:r><w:t>only paragraph</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>',
    );
    expect(scanDocumentXml(xml).bodyFeatures).not.toContain("multi-section");
  });

  test("CJK text is tagged 'cjk'", () => {
    const paragraph = scanDocumentXml(wrapBody("<w:p><w:r><w:t>你好世界</w:t></w:r></w:p>"))
      .paragraphs[0];
    expect(paragraph?.features).toEqual(["cjk"]);
    expect(paragraph?.normText).toBe("你好世界");
  });

  test("XML entities (named + numeric decimal + numeric hex) decode inside w:t", () => {
    const xml = wrapBody(
      '<w:p><w:r><w:t xml:space="preserve">Tom &amp; Jerry &lt;3 &quot;quote&quot; &#65;&#x42;</w:t></w:r></w:p>',
    );
    const paragraph = scanDocumentXml(xml).paragraphs[0];
    expect(paragraph?.normText).toBe('Tom & Jerry <3 "quote" AB');
  });

  test("self-closing empty <w:p/> paragraphs are kept (with depth-derived features only)", () => {
    const bare = scanDocumentXml(wrapBody("<w:p/>")).paragraphs;
    expect(bare).toEqual([{ normText: "", features: [] }]);

    const inTable = scanDocumentXml(
      wrapBody("<w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>"),
    ).paragraphs;
    expect(inTable).toEqual([{ normText: "", features: ["table"] }]);
  });

  test("multiple paragraphs are scanned in document order", () => {
    const xml = wrapBody(
      "<w:p><w:r><w:t>First</w:t></w:r></w:p>" +
        "<w:p/>" +
        "<w:p><w:r><w:t>Third</w:t></w:r></w:p>",
    );
    const { paragraphs } = scanDocumentXml(xml);
    expect(paragraphs.map((p) => p.normText)).toEqual(["First", "", "Third"]);
  });

  test("multi-column and landscape doc-level tags", () => {
    const xml = wrapBody(
      '<w:p><w:r><w:t>x</w:t></w:r></w:p><w:sectPr><w:pgSz w:orient="landscape" w:w="15840" w:h="12240"/><w:cols w:num="3"/></w:sectPr>',
    );
    const { bodyFeatures } = scanDocumentXml(xml);
    expect(bodyFeatures).toContain("multi-column");
    expect(bodyFeatures).toContain("landscape");
  });

  test("w:cols w:num=1 does not count as multi-column", () => {
    const xml = wrapBody(
      '<w:p><w:r><w:t>x</w:t></w:r></w:p><w:sectPr><w:cols w:num="1"/></w:sectPr>',
    );
    expect(scanDocumentXml(xml).bodyFeatures).not.toContain("multi-column");
  });
});

describe("attributeDivergences", () => {
  const baseResult = (divergences: Divergence[]): ParityResult => ({
    file: "/docs/sample.docx",
    score: 0.9,
    referencePages: 1,
    folioPages: 1,
    totalReferenceLines: 10,
    matchedLines: 9,
    medianYOffsetPt: 0,
    divergences,
  });

  const paragraph = (normText: string, features: string[]): ParagraphFeatures => ({
    normText,
    features,
  });

  test("matches by substring: divergence text contained in a longer paragraph normText", () => {
    const doc: DocFeatures = {
      paragraphs: [paragraph("This is a long paragraph with a tab stop in it", ["tab-stops"])],
      docFeatures: ["headers"],
    };
    const result = baseResult([{ kind: "missing-line", page: 1, text: "a tab stop" }]);
    const attributed = attributeDivergences(result, doc);
    expect(attributed.attributed[0]?.features).toEqual(["tab-stops"]);
  });

  test("matches by substring: paragraph normText contained in a longer divergence text", () => {
    const doc: DocFeatures = {
      paragraphs: [paragraph("Short cell", ["table"])],
      docFeatures: [],
    };
    const result = baseResult([
      { kind: "extra-line", page: 2, text: "Short cell (folio rendered extra context)" },
    ]);
    const attributed = attributeDivergences(result, doc);
    expect(attributed.attributed[0]?.features).toEqual(["table"]);
  });

  test("falls back to best textSimilarity above the 0.6 threshold when no substring matches", () => {
    const doc: DocFeatures = {
      paragraphs: [
        paragraph("The quick brown fox jumps over", ["justify"]),
        paragraph("Completely unrelated content here", ["numbering"]),
      ],
      docFeatures: [],
    };
    // One character different from the first paragraph: no substring match,
    // but similarity is high.
    const result = baseResult([
      { kind: "missing-line", page: 1, text: "The quick brown fox jumps ovex" },
    ]);
    const attributed = attributeDivergences(result, doc);
    expect(attributed.attributed[0]?.features).toEqual(["justify"]);
  });

  test("falls back to doc:-prefixed docFeatures when nothing matches", () => {
    const doc: DocFeatures = {
      paragraphs: [paragraph("Nothing like it", ["numbering"])],
      docFeatures: ["footnotes", "landscape"],
    };
    const result = baseResult([
      { kind: "missing-line", page: 1, text: "Totally different text entirely" },
    ]);
    const attributed = attributeDivergences(result, doc);
    expect(attributed.attributed[0]?.features).toEqual(["doc:footnotes", "doc:landscape"]);
  });

  test("falls back to ['doc:unattributed'] when doc has no docFeatures either", () => {
    const doc: DocFeatures = { paragraphs: [], docFeatures: [] };
    const result = baseResult([{ kind: "missing-line", page: 1, text: "Anything" }]);
    const attributed = attributeDivergences(result, doc);
    expect(attributed.attributed[0]?.features).toEqual(["doc:unattributed"]);
  });

  test("page-count divergences (no text) always attribute via docFeatures", () => {
    const doc: DocFeatures = {
      paragraphs: [paragraph("Whatever", ["table"])],
      docFeatures: ["multi-section"],
    };
    const result = baseResult([{ kind: "page-count", reference: 3, folio: 4 }]);
    const attributed = attributeDivergences(result, doc);
    expect(attributed.attributed[0]?.features).toEqual(["doc:multi-section"]);
  });

  test("short strings (< 4 chars) never match by substring, avoiding junk matches", () => {
    // "ab" is too short (< MIN_MATCH_LEN) for the substring step even though
    // it is trivially a substring of "abcdef"; the similarity fallback also
    // fails ("ab" vs "abcdef" scores well below 0.6), so this falls through
    // to docFeatures.
    const doc: DocFeatures = {
      paragraphs: [paragraph("abcdef", ["table"])],
      docFeatures: ["headers"],
    };
    const result = baseResult([{ kind: "missing-line", page: 1, text: "ab" }]);
    const attributed = attributeDivergences(result, doc);
    expect(attributed.attributed[0]?.features).toEqual(["doc:headers"]);
  });

  test("copies ParityResult fields through and attaches docFeatures", () => {
    const doc: DocFeatures = { paragraphs: [], docFeatures: ["landscape"] };
    const result = baseResult([{ kind: "page-count", reference: 1, folio: 1 }]);
    const attributed = attributeDivergences(result, doc);
    expect(attributed.file).toBe(result.file);
    expect(attributed.score).toBe(result.score);
    expect(attributed.docFeatures).toEqual(["landscape"]);
  });
});

describe("clusterCorpus", () => {
  const paragraph = (normText: string, features: string[]): ParagraphFeatures => ({
    normText,
    features,
  });

  const makeResult = (
    file: string,
    docFeatures: string[],
    attributed: FeatureAttributedResult["attributed"],
  ): FeatureAttributedResult => ({
    file,
    score: 0.9,
    referencePages: 1,
    folioPages: 1,
    totalReferenceLines: 10,
    matchedLines: 9,
    medianYOffsetPt: 0,
    divergences: attributed.map((a) => a.divergence),
    attributed,
    docFeatures,
  });

  const yDrift = (page: number, text: string, residualPt: number, features: string[]) => ({
    divergence: { kind: "y-drift" as const, page, text, residualPt },
    features,
  });

  const xDrift = (page: number, text: string, deltaPt: number, features: string[]) => ({
    divergence: { kind: "x-drift" as const, page, text, deltaPt },
    features,
  });

  test("a feature rare in the corpus but common among one divergence kind gets high lift and ranks first", () => {
    // Baseline pool: 40 paragraphs total, only 2 carry "table" (5% prevalence).
    const corpusParagraphs: ParagraphFeatures[][] = [
      Array.from({ length: 19 }, (_, i) => paragraph(`plain paragraph ${i}`, [])),
      [paragraph("table cell one", ["table"]), paragraph("table cell two", ["table"])],
      Array.from({ length: 19 }, (_, i) => paragraph(`other paragraph ${i}`, ["justify"])),
    ];

    const results: FeatureAttributedResult[] = [
      makeResult(
        "/a.docx",
        [],
        [yDrift(1, "table cell one", 4, ["table"]), yDrift(1, "table cell two", 5, ["table"])],
      ),
      makeResult("/b.docx", [], [yDrift(1, "table cell three", 6, ["table"])]),
      makeResult("/c.docx", [], [yDrift(1, "unrelated line", 1, ["justify"])]),
      makeResult(
        "/d.docx",
        [],
        [xDrift(1, "some line", 1, ["justify"]), xDrift(1, "another line", 1.2, ["justify"])],
      ),
      makeResult("/e.docx", [], [xDrift(1, "yet another", 0.9, ["justify"])]),
      makeResult("/f.docx", [], []),
    ];

    const clusters = clusterCorpus(results, corpusParagraphs);
    const tableYDrift = clusters.find((c) => c.kind === "y-drift" && c.feature === "table");
    expect(tableYDrift).toBeDefined();
    expect(tableYDrift?.count).toBe(3);
    expect(tableYDrift?.docs).toEqual(["/a.docx", "/b.docx"]);
    // observedRate = 3/4 (3 of the 4 y-drift divergences carry "table");
    // baseline = 2/40 = 0.05; lift = 0.75 / 0.05 = 15.
    expect(tableYDrift?.lift).toBeCloseTo(15, 5);
    expect(tableYDrift?.meanMagnitudePt).toBeCloseTo((4 + 5 + 6) / 3, 5);

    // Ranked first: lift * sqrt(count) for "table"/y-drift is far larger than
    // the "justify" clusters, whose observed rate roughly matches baseline.
    expect(clusters[0]).toBe(tableYDrift);
  });

  test("count filtering: a (kind, feature) pair seen only once is dropped when the corpus is not small", () => {
    const corpusParagraphs: ParagraphFeatures[][] = [
      Array.from({ length: 10 }, (_, i) => paragraph(`p${i}`, ["indent"])),
    ];
    const results: FeatureAttributedResult[] = [
      makeResult("/a.docx", [], [yDrift(1, "x", 1, ["indent"])]),
      makeResult("/b.docx", [], []),
      makeResult("/c.docx", [], []),
      makeResult("/d.docx", [], []),
      makeResult("/e.docx", [], []),
    ];
    const clusters = clusterCorpus(results, corpusParagraphs);
    expect(clusters.find((c) => c.feature === "indent")).toBeUndefined();
  });

  test("small corpus (< 5 docs) allows count === 1 clusters through", () => {
    const corpusParagraphs: ParagraphFeatures[][] = [[paragraph("p", ["indent"])]];
    const results: FeatureAttributedResult[] = [
      makeResult("/a.docx", [], [yDrift(1, "x", 1, ["indent"])]),
      makeResult("/b.docx", [], []),
    ];
    const clusters = clusterCorpus(results, corpusParagraphs);
    const cluster = clusters.find((c) => c.feature === "indent");
    expect(cluster?.count).toBe(1);
  });

  test("examples are capped at 5 even when count is higher", () => {
    const corpusParagraphs: ParagraphFeatures[][] = [
      Array.from({ length: 10 }, (_, i) => paragraph(`p${i}`, ["indent"])),
    ];
    const attributed = Array.from({ length: 7 }, (_, i) => yDrift(1, `line ${i}`, 1, ["indent"]));
    const results: FeatureAttributedResult[] = [makeResult("/a.docx", [], attributed)];
    const clusters = clusterCorpus(results, corpusParagraphs);
    const cluster = clusters.find((c) => c.feature === "indent");
    expect(cluster?.count).toBe(7);
    expect(cluster?.examples).toHaveLength(5);
  });

  test("doc:-prefixed features are baselined against docs (results), not paragraphs", () => {
    // No paragraph anywhere carries "landscape" (it's a doc-level tag), so
    // the paragraph-baseline would floor to MIN_BASELINE_PREVALENCE and
    // produce an enormous, meaningless lift. The doc-baseline instead sees
    // 1 of 2 docs (50%) carrying "doc:landscape".
    const corpusParagraphs: ParagraphFeatures[][] = [
      Array.from({ length: 10 }, (_, i) => paragraph(`p${i}`, [])),
      Array.from({ length: 10 }, (_, i) => paragraph(`q${i}`, [])),
    ];
    const results: FeatureAttributedResult[] = [
      makeResult(
        "/a.docx",
        ["landscape"],
        [
          {
            divergence: { kind: "page-count", reference: 1, folio: 2 },
            features: ["doc:landscape"],
          },
          {
            divergence: { kind: "page-count", reference: 1, folio: 2 },
            features: ["doc:landscape"],
          },
        ],
      ),
      makeResult("/b.docx", [], []),
    ];
    const clusters = clusterCorpus(results, corpusParagraphs);
    const cluster = clusters.find((c) => c.feature === "doc:landscape");
    expect(cluster).toBeDefined();
    // observedRate = 2/2 = 1; baseline = 1/2 = 0.5; lift = 2.
    expect(cluster?.lift).toBeCloseTo(2, 5);
  });
});

describe("computeFontSubstitutionTags", () => {
  test("PostScript style decorations satisfy the requested family", () => {
    expect(
      computeFontSubstitutionTags(
        ["Arial", "Calibri", "Times New Roman"],
        ["ArialMT", "Calibri-BoldItalic", "TimesNewRomanPSMT"],
      ),
    ).toEqual([]);
  });

  test("a different family sharing a prefix is a substitution", () => {
    // "Interstate-Bold" must NOT satisfy a request for "Inter".
    expect(computeFontSubstitutionTags(["Inter"], ["Interstate-Bold"])).toEqual([
      "font-substituted:inter",
    ]);
  });

  test("subset prefixes are stripped and duplicates deduplicated", () => {
    expect(
      computeFontSubstitutionTags(["Calibri", "calibri", "Missing Font"], ["ABCDEF+Calibri-Bold"]),
    ).toEqual(["font-substituted:missingfont"]);
  });
});

const fontGeom = (source: "word" | "folio", lines: Array<[string, string]>): DocGeom => ({
  source,
  file: "/font-test.docx",
  pages: [
    {
      number: 1,
      widthPt: 600,
      heightPt: 800,
      lines: lines.map(([text, fontName], index) => ({
        text,
        normText: text,
        xPt: 0,
        yPt: index * 12,
        widthPt: 100,
        heightPt: 12,
        fontName,
        region: "body",
      })),
    },
  ],
  meta: {},
});

describe("assessFontEnvironment", () => {
  test("recognizes equivalent PDF and CSS family names", () => {
    expect(fontFamiliesMatch("ArialMT", "Arial")).toBe(true);
    expect(fontFamiliesMatch("ABCDEF+Calibri-BoldItalic", "Calibri")).toBe(true);
    expect(fontFamiliesMatch("Interstate-Bold", "Inter")).toBe(false);
  });

  test("accepts a shared substituted family", () => {
    const assessment = assessFontEnvironment(
      ["Times New Roman"],
      fontGeom("word", [["Shared line", "Tinos-Regular"]]),
      fontGeom("folio", [["Shared line", "Tinos"]]),
    );

    expect(assessment).toEqual({
      status: "shared-substitution",
      tags: ["font-shared:timesnewroman"],
      comparedLines: 1,
      matchingLines: 1,
    });
  });

  test("rejects a cross-renderer family mismatch", () => {
    const assessment = assessFontEnvironment(
      ["Times New Roman"],
      fontGeom("word", [["Same text", "LiberationSerif"]]),
      fontGeom("folio", [["Same text", "Tinos"]]),
    );

    expect(assessment).toEqual({
      status: "mismatch",
      tags: ["font-renderer-mismatch", "font-substituted:timesnewroman"],
      comparedLines: 1,
      matchingLines: 0,
    });
  });

  test("rejects a renderer mismatch even when the reference used the requested font", () => {
    const assessment = assessFontEnvironment(
      ["Arial"],
      fontGeom("word", [["Same text", "ArialMT"]]),
      fontGeom("folio", [["Same text", "Helvetica"]]),
    );

    expect(assessment).toEqual({
      status: "mismatch",
      tags: ["font-renderer-mismatch"],
      comparedLines: 1,
      matchingLines: 0,
    });
  });

  test("rejects matching family names with persistently different metrics", () => {
    const lines = Array.from({ length: 8 }, (_, index): [string, string] => [
      `Line ${index}`,
      "Arial",
    ]);
    const word = fontGeom("word", lines);
    const folio = fontGeom("folio", lines);
    for (const line of folio.pages[0]?.lines ?? []) {
      line.widthPt = 110;
    }

    const assessment = assessFontEnvironment(["Arial"], word, folio);

    expect(assessment).toEqual({
      status: "mismatch",
      tags: ["font-renderer-metric-mismatch"],
      comparedLines: 8,
      matchingLines: 8,
    });
  });

  test("detects a repeated metric cluster hidden by otherwise stable lines", () => {
    const lines = Array.from({ length: 32 }, (_, index): [string, string] => [
      `Line ${index}`,
      "Arial",
    ]);
    const word = fontGeom("word", lines);
    const folio = fontGeom("folio", lines);
    for (const line of folio.pages[0]?.lines.slice(0, 8) ?? []) {
      line.widthPt = 110;
    }

    const assessment = assessFontEnvironment(["Arial"], word, folio);

    expect(assessment).toEqual({
      status: "mismatch",
      tags: ["font-renderer-metric-mismatch"],
      comparedLines: 32,
      matchingLines: 32,
    });
  });

  test("reports unverified when no font-bearing text can be paired", () => {
    const assessment = assessFontEnvironment(
      ["Arial"],
      fontGeom("word", [["Word only", "ArialMT"]]),
      fontGeom("folio", [["Folio only", "Arial"]]),
    );

    expect(assessment.status).toBe("unverified");
    expect(assessment.tags).toEqual(["font-parity-unverified"]);
  });
});
