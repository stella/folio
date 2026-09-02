import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { compileLegalSourceToDocx, compileLegalSourceToDocument, parseLegalSource } from "./index";

const SAMPLE_SOURCE = `@doc kind=agreement locale=en-GB numbering=legal page=A4
@title Mutual Non-Disclosure Agreement

@preamble
This Agreement is made between Alpha Ltd and Beta s.r.o.

@recital
The parties wish to exchange confidential information for the Purpose.

@clause 1. Definitions
"Confidential Information" means all non-public information disclosed by either party.

@subclause Permitted Disclosure
A party may disclose Confidential Information to its professional advisers.

@table
| Item | Responsible Party | Status |
| --- | --- | --- |
| Board approval | Alpha Ltd | Open |
| Financing consent | Beta s.r.o. |

@signatures
Alpha Ltd
Beta s.r.o.
`;

describe("Stella Legal Source", () => {
  test("parses compact legal directives and applies deterministic autofixes", () => {
    const result = parseLegalSource(SAMPLE_SOURCE);

    expect(result.diagnostics.filter((d) => d.severity === "error")).toEqual([]);
    expect(result.draft.meta.title).toBe("Mutual Non-Disclosure Agreement");
    expect(result.fixes.map((fix) => fix.code)).toContain("manual-numbering-stripped");
    expect(result.fixes.map((fix) => fix.code)).toContain("table-row-width-normalized");

    const definitions = result.draft.blocks.find(
      (block) => block.type === "clause" && block.heading === "Definitions",
    );
    expect(definitions).toBeDefined();
  });

  test("keeps a pipe-table row that omits the trailing pipe", () => {
    const source = `@doc kind=agreement locale=en-GB numbering=legal page=A4
@title T

@table
| Item | Party |
| --- | --- |
| Approval | Alpha Ltd |
| Consent | Beta Ltd
| Funding | Gamma Ltd |
`;
    const result = parseLegalSource(source);
    const table = result.draft.blocks.find((block) => block.type === "table");
    if (table?.type !== "table") {
      throw new Error("expected a table block");
    }
    expect(table.table.rows).toHaveLength(3);
    const cells = table.table.rows.flat();
    expect(cells).toContain("Consent");
    expect(cells).toContain("Beta Ltd");
  });

  test("parses quoted doc attributes and reports invalid known values", () => {
    const result = parseLegalSource(
      [
        '@doc kind=contract numbering=bogus page=Legal orientation=sideways title="Share Purchase Agreement"',
        "@paragraph",
        "Body text.",
      ].join("\n"),
    );

    expect(result.draft.meta.title).toBe("Share Purchase Agreement");
    expect(result.diagnostics).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          code: "invalid-doc-attribute",
          message: 'Invalid @doc kind "contract".',
        }),
        expect.objectContaining({
          code: "invalid-doc-attribute",
          message: 'Invalid @doc numbering "bogus".',
        }),
        expect.objectContaining({
          code: "invalid-doc-attribute",
          message: 'Invalid @doc page "Legal".',
        }),
        expect.objectContaining({
          code: "invalid-doc-attribute",
          message: 'Invalid @doc orientation "sideways".',
        }),
      ]),
    );
  });

  test("parses structured signature party fields", () => {
    const result = parseLegalSource(
      [
        '@doc title="Execution Version"',
        "@signatures",
        "Party: Alpha Ltd",
        "By: Jane Doe",
        "Title: Chief Executive Officer",
        "Beta s.r.o.",
        "Name: Jan Novak",
        "Title: Jednatel",
      ].join("\n"),
    );

    const signatures = result.draft.blocks.find((block) => block.type === "signatures");

    expect(signatures).toEqual({
      type: "signatures",
      parties: [
        {
          name: "Alpha Ltd",
          signatory: "Jane Doe",
          title: "Chief Executive Officer",
        },
        { name: "Beta s.r.o.", signatory: "Jan Novak", title: "Jednatel" },
      ],
    });
  });

  test("strips manual ordered-list markers before applying DOCX numbering", () => {
    const result = parseLegalSource(
      [
        '@doc title="Numbered List"',
        "@list ordered",
        "1. First item",
        "2) Second item",
        "3.1 Third item",
        "2024 budget approval",
        "30 day notice period",
      ].join("\n"),
    );

    const list = result.draft.blocks.find((block) => block.type === "list");

    expect(list).toEqual({
      type: "list",
      ordered: true,
      items: [
        "First item",
        "Second item",
        "Third item",
        "2024 budget approval",
        "30 day notice period",
      ],
    });
  });

  test("keeps distinct non-Latin clauses when removing duplicate titles", () => {
    const result = parseLegalSource(
      ['@doc title="服务协议"', "@clause 保密义务", "双方应保护机密信息。"].join("\n"),
    );

    const clauses = result.draft.blocks.filter((block) => block.type === "clause");

    expect(clauses.map((clause) => clause.heading)).toEqual(["保密义务"]);
    expect(result.fixes.map((fix) => fix.code)).not.toContain("duplicate-title-clause-removed");
  });

  test("removes only a leading duplicate-title clause", () => {
    const result = parseLegalSource(
      [
        '@doc title="Master Services Agreement"',
        "@clause Master Services Agreement",
        "Introductory duplicate title.",
        "@clause Background",
        "The parties have prior dealings.",
        "@clause Master Services Agreement",
        "This later clause is intentional.",
      ].join("\n"),
    );

    const clauses = result.draft.blocks.filter((block) => block.type === "clause");

    expect(clauses.map((clause) => clause.heading)).toEqual([
      "Background",
      "Master Services Agreement",
    ]);
    expect(
      result.fixes.filter((fix) => fix.code === "duplicate-title-clause-removed"),
    ).toHaveLength(1);
  });

  test("does not strip legitimate one-letter clause heading words", () => {
    const result = parseLegalSource(
      [
        '@doc title="Agreement"',
        "@clause A Party's Obligations",
        "Each party must comply.",
        "@clause A. Definitions",
        "Defined terms apply.",
      ].join("\n"),
    );

    const clauses = result.draft.blocks.filter((block) => block.type === "clause");

    expect(clauses.map((clause) => clause.heading)).toEqual([
      "A Party's Obligations",
      "Definitions",
    ]);
  });

  test("compiles to the canonical document model with legal numbering", () => {
    const result = compileLegalSourceToDocument(SAMPLE_SOURCE);

    expect(result.status).toBe("ok");
    if (result.status !== "ok") {
      return;
    }

    const paragraphs = result.document.package.document.content.filter(
      (block) => block.type === "paragraph",
    );
    const numbered = paragraphs.filter((paragraph) => paragraph.formatting?.numPr !== undefined);

    expect(numbered.length).toBeGreaterThanOrEqual(2);
    expect(result.document.package.numbering?.abstractNums.at(0)?.levels).toHaveLength(5);
  });

  test("compiles numbering=none without numbering definitions or paragraph numPr", async () => {
    const source = [
      '@doc numbering=none title="Unnumbered Memo"',
      "@clause Background",
      "The parties agree the document should remain unnumbered.",
      "@list",
      "- First item",
      "- Second item",
    ].join("\n");
    const result = compileLegalSourceToDocument(source);

    expect(result.status).toBe("ok");
    if (result.status !== "ok") {
      return;
    }

    expect(result.document.package.numbering).toBeUndefined();
    const paragraphs = result.document.package.document.content.filter(
      (block) => block.type === "paragraph",
    );
    expect(paragraphs.every((paragraph) => paragraph.formatting?.numPr === undefined)).toBe(true);

    const docxResult = await compileLegalSourceToDocx(source);
    expect(docxResult.status).toBe("ok");
    if (docxResult.status !== "ok") {
      return;
    }

    const zip = await JSZip.loadAsync(docxResult.buffer);
    expect(zip.file("word/numbering.xml")).toBeNull();
    const contentTypesXml = await zip.file("[Content_Types].xml")?.async("string");
    const documentRelsXml = await zip.file("word/_rels/document.xml.rels")?.async("string");
    const documentXml = await zip.file("word/document.xml")?.async("string");
    expect(contentTypesXml).not.toContain("numbering.xml");
    expect(documentRelsXml).not.toContain("numbering.xml");
    expect(documentXml).not.toContain("<w:numPr>");
  });

  test("compiles checklist numbering as checkbox list items without legal clause numbering", () => {
    const source = [
      '@doc numbering=checklist title="Closing Checklist"',
      "@clause Before Completion",
      "Confirm each condition.",
      "@list",
      "- Board approval received",
      "- Funds flow approved",
    ].join("\n");
    const result = compileLegalSourceToDocument(source);

    expect(result.status).toBe("ok");
    if (result.status !== "ok") {
      return;
    }

    const paragraphs = result.document.package.document.content.filter(
      (block) => block.type === "paragraph",
    );
    const clause = paragraphs.find(
      (paragraph) => paragraph.formatting?.styleId === "ClauseHeading1",
    );
    const listItems = paragraphs.filter(
      (paragraph) => paragraph.formatting?.styleId === "ListParagraph",
    );

    expect(clause?.formatting?.numPr).toBeUndefined();
    expect(listItems.map((paragraph) => paragraph.formatting?.numPr)).toEqual([
      { numId: 3, ilvl: 0 },
      { numId: 3, ilvl: 0 },
    ]);
    expect(result.document.package.numbering?.abstractNums).toEqual([
      expect.objectContaining({
        abstractNumId: 3,
        levels: [expect.objectContaining({ lvlText: "☐" })],
      }),
    ]);
  });

  test("serializes a valid DOCX package without the docx npm model", async () => {
    const result = await compileLegalSourceToDocx(SAMPLE_SOURCE);

    expect(result.status).toBe("ok");
    if (result.status !== "ok") {
      return;
    }

    expect(result.buffer).toBeInstanceOf(Buffer);
    const zip = await JSZip.loadAsync(result.buffer);
    expect(zip.file("word/document.xml")).toBeTruthy();
    expect(zip.file("word/styles.xml")).toBeTruthy();
    expect(zip.file("word/numbering.xml")).toBeTruthy();

    const documentXml = await zip.file("word/document.xml")?.async("string");
    // Title text is stored in original case; the Title style's allCaps renders it
    // uppercase in Word (a manual toUpperCase would corrupt non-English casing).
    expect(documentXml).toContain("Mutual Non-Disclosure Agreement");
    expect(documentXml).not.toContain("MUTUAL NON-DISCLOSURE AGREEMENT");
    expect(documentXml).toContain("<w:numPr>");
    expect(documentXml).toContain("<w:tbl>");
    expect(documentXml).toContain("<w:tblGrid>");

    const numberingXml = await zip.file("word/numbering.xml")?.async("string");
    expect(numberingXml).toContain("<w:nsid");
    expect(numberingXml).toContain("<w:lvlJc");
    // CT_Lvl child order (ECMA-376): isLgl and suff must precede lvlText.
    expect(numberingXml).toContain('<w:isLgl/><w:suff w:val="tab"/>');
    expect(numberingXml).toContain('<w:suff w:val="tab"/><w:lvlText');
  });
});

const bodyParagraphs = (result: ReturnType<typeof compileLegalSourceToDocument>) =>
  result.status === "ok"
    ? result.document.package.document.content.flatMap((block) =>
        block.type === "paragraph" && block.formatting?.styleId === "BodyText" ? [block] : [],
      )
    : [];

const runsOf = (paragraph: { content: { type: string }[] }) =>
  paragraph.content.flatMap((node) => {
    if (node.type === "run" && "formatting" in node && "content" in node) {
      return [node];
    }
    if (node.type === "hyperlink" && "children" in node && Array.isArray(node.children)) {
      return node.children;
    }
    return [];
  });

describe("Stella Legal Source — markdown bodies", () => {
  test("renders inline markdown in a clause body as formatted runs", () => {
    const result = compileLegalSourceToDocument(
      [
        '@doc title="Inline"',
        "@clause Payment",
        "The **Buyer** pays *promptly*, see [the schedule](https://example.com/s).",
      ].join("\n"),
    );
    const [body] = bodyParagraphs(result);
    expect(body).toBeDefined();
    const hyperlink = body?.content.find((node) => node.type === "hyperlink");
    expect(hyperlink?.type === "hyperlink" ? hyperlink.href : undefined).toBe(
      "https://example.com/s",
    );
    expect(JSON.stringify(body)).toContain('"bold":true');
    expect(JSON.stringify(body)).toContain('"italic":true');
    expect(JSON.stringify(body)).not.toContain("**");
  });

  test("keeps [[placeholders]] highlighted inside emphasised text", () => {
    const result = compileLegalSourceToDocument(
      ['@doc title="Placeholders"', "@paragraph", "**[[Party]]** shall pay [[Amount]]."].join("\n"),
    );
    const [body] = bodyParagraphs(result);
    const highlighted = (body ? runsOf(body) : []).filter(
      (run) => "formatting" in run && run.formatting?.highlight === "yellow",
    );
    expect(highlighted).toHaveLength(2);
  });

  test("a markdown list inside a clause body becomes a list block in place", () => {
    const result = parseLegalSource(
      [
        '@doc title="Lists"',
        "@clause Duties",
        "The Supplier shall:",
        "- deliver on time",
        "- report monthly",
        "",
        "Both duties are ongoing.",
      ].join("\n"),
    );
    expect(result.draft.blocks).toEqual([
      { type: "clause", level: 1, heading: "Duties", paragraphs: ["The Supplier shall:"] },
      { type: "list", ordered: false, items: ["deliver on time", "report monthly"] },
      { type: "paragraph", paragraphs: ["Both duties are ongoing."] },
    ]);
  });

  test("a GFM pipe table outside @table is a table block", () => {
    const result = parseLegalSource(
      ['@doc title="Tables"', "| Item | Owner |", "| --- | --- |", "| Approval | Alpha |"].join(
        "\n",
      ),
    );
    expect(result.draft.blocks).toEqual([
      { type: "table", table: { headers: ["Item", "Owner"], rows: [["Approval", "Alpha"]] } },
    ]);
  });

  test("a directive line never merges into the paragraph above it", () => {
    const result = parseLegalSource(
      ['@doc title="T"', "@clause First", "Body line", "@clause Next", "More"].join("\n"),
    );
    expect(result.draft.blocks).toEqual([
      { type: "clause", level: 1, heading: "First", paragraphs: ["Body line"] },
      { type: "clause", level: 1, heading: "Next", paragraphs: ["More"] },
    ]);
  });

  test("soft-wrapped lines stay one paragraph", () => {
    const result = parseLegalSource(
      [
        '@doc title="Wrap"',
        "@paragraph",
        "This sentence",
        "continues here.",
        "",
        "Second one.",
      ].join("\n"),
    );
    expect(result.draft.blocks).toEqual([
      { type: "paragraph", paragraphs: ["This sentence continues here.", "Second one."] },
    ]);
  });

  test("reports an unknown directive as the author spelled it", () => {
    const result = parseLegalSource(['@doc title="T"', "@Whereas the parties agree"].join("\n"));
    expect(result.diagnostics).toContainEqual(
      expect.objectContaining({
        code: "unknown-directive",
        message: 'Unknown legal directive "@Whereas".',
        line: 2,
      }),
    );
  });

  test("a blank line before a directive never swallows it", () => {
    const result = parseLegalSource(
      [
        "",
        "@doc title=T",
        "",
        "@title T",
        "",
        "@clause A",
        "body",
        "",
        "@pagebreak",
        "",
        "@clause B",
        "more",
      ].join("\n"),
    );
    expect(result.draft.blocks).toEqual([
      { type: "title", text: "T" },
      { type: "clause", level: 1, heading: "A", paragraphs: ["body"] },
      { type: "pageBreak" },
      { type: "clause", level: 1, heading: "B", paragraphs: ["more"] },
    ]);
  });

  test("fenced code stays literal through the inline renderer", () => {
    const result = compileLegalSourceToDocument(
      ['@doc title="Code"', "@paragraph", "```", "**not bold** [[not a placeholder]]", "```"].join(
        "\n",
      ),
    );
    const [body] = bodyParagraphs(result);
    expect(body).toBeDefined();
    const runs = body ? runsOf(body) : [];
    // Escaped characters come back as separate runs; the joined text is
    // what matters.
    expect(
      runs
        .map((run) =>
          "content" in run
            ? run.content.map((item) => (item.type === "text" ? item.text : "")).join("")
            : "",
        )
        .join(""),
    ).toBe("**not bold** [[not a placeholder]]");
    expect(JSON.stringify(body)).not.toContain('"bold":true');
    expect(JSON.stringify(body)).not.toContain('"highlight"');
  });

  test("a list or quote directly followed by a directive does not swallow it", () => {
    const result = parseLegalSource(
      [
        '@doc title="Boundaries"',
        "@clause Duties",
        "- deliver",
        "- report",
        "@clause Fees",
        "> quoted intro",
        "@subclause Rates",
        "Hourly.",
      ].join("\n"),
    );
    expect(result.draft.blocks).toEqual([
      { type: "clause", level: 1, heading: "Duties", paragraphs: [] },
      { type: "list", ordered: false, items: ["deliver", "report"] },
      { type: "clause", level: 1, heading: "Fees", paragraphs: ["quoted intro"] },
      { type: "clause", level: 2, heading: "Rates", paragraphs: ["Hourly."] },
    ]);
  });

  test("diagnostics keep the author's line numbers after boundary insertion", () => {
    const result = parseLegalSource(
      ['@doc title="Lines"', "@list", "- one", "@bogus directive", "@clause A"].join("\n"),
    );
    expect(result.diagnostics).toContainEqual(
      expect.objectContaining({ code: "unknown-directive", line: 4 }),
    );
  });

  test("tolerates closing directives a model invents", () => {
    const result = parseLegalSource(
      [
        '@doc title="Closers"',
        "@list",
        "- one",
        "- two",
        "@endlist",
        "@clause A",
        "body",
        "@end",
      ].join("\n"),
    );
    expect(result.diagnostics).toEqual([]);
    expect(result.fixes.filter((fix) => fix.code === "closing-directive-ignored")).toHaveLength(2);
    expect(result.draft.blocks).toEqual([
      { type: "list", ordered: false, items: ["one", "two"] },
      { type: "clause", level: 1, heading: "A", paragraphs: ["body"] },
    ]);
  });

  test("warns about a body paragraph that is bold end to end instead of rewriting it", () => {
    const result = compileLegalSourceToDocument(
      ['@doc title="Bold"', "@paragraph", "**Whole sentence carried over from chat.**"].join("\n"),
    );
    expect(result.status).toBe("ok");
    if (result.status !== "ok") {
      return;
    }
    expect(result.warnings.map((warning) => warning.code)).toEqual(["whole-paragraph-emphasis"]);
    expect(JSON.stringify(bodyParagraphs(result))).toContain('"bold":true');
  });

  test("signature fields stay literal, never markdown, but keep placeholder highlights", () => {
    const result = compileLegalSourceToDocument(
      [
        '@doc title="Sign"',
        "@signatures",
        "Party: A_B_C Ltd",
        "By: *Jane*",
        "Party: [[Seller legal name]]",
        "By: [[name of signatory]]",
      ].join("\n"),
    );
    expect(result.status).toBe("ok");
    if (result.status !== "ok") {
      return;
    }
    const serialized = JSON.stringify(result.document.package.document.content);
    expect(serialized).toContain("A_B_C Ltd");
    expect(serialized).toContain("*Jane*");
    expect(serialized).not.toContain("[[");
    expect(serialized).toContain(
      JSON.stringify({ text: "Seller legal name", preserveSpace: true }).slice(1, -1),
    );
    const highlighted = (serialized.match(/"highlight":"yellow"/gu) ?? []).length;
    expect(highlighted).toBe(2);
  });
});
