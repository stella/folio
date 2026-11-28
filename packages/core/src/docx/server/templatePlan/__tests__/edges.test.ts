import { describe, expect, test } from "bun:test";

import { applyTemplatePlan } from "../apply";
import { analyzeDocxBlocks } from "../blocks";
import { inspectDocxTemplate } from "../inspect";
import { innerRangeOf, readRawPart } from "../xml";
import {
  buildMinimalDocx,
  buildTemplate,
  numPr,
  p,
  r,
  readPart,
  replacePart,
  TEMPLATE_IDS,
} from "./fixtures";
import {
  applyFor,
  buildSourceWith,
  planFor,
  unwrap,
  verifyFor,
} from "./planning";

describe("innerRangeOf", () => {
  test("finds the content after an opening tag whose attribute holds '>'", () => {
    const xml =
      '<w:root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p w:x="a>b"><w:r/></w:p><w:q/></w:root>';
    const part = readRawPart("x", xml);
    const [first, second] = part?.root.elements ?? [];
    if (first === undefined || second === undefined) {
      throw new Error("fixture did not parse");
    }
    const inner = innerRangeOf(xml, first);

    expect(inner === null ? null : xml.slice(inner.start, inner.end)).toBe(
      "<w:r/>",
    );
    expect(innerRangeOf(xml, second)).toBeNull();
  });
});

describe("applyTemplatePlan edge content", () => {
  test("carries a field that numbers itself as the field it is", async () => {
    const template = await buildTemplate();
    const number =
      '<w:r><w:rPr><w:b/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:rPr><w:b/></w:rPr><w:instrText xml:space="preserve"> AUTONUMLGL  \\* Arabic \\e </w:instrText></w:r><w:r><w:rPr><w:b/></w:rPr><w:fldChar w:fldCharType="end"/></w:r>';
    const source = await buildMinimalDocx(
      [
        p(`${r("Annex ")}${number}${r(" - Summary")}`),
        p(`${r("Annex ")}<w:fldSimple w:instr=" AUTONUM "/>${r(" - Lists")}`),
      ].join(""),
    );
    // The manifest says where the numbers stand that its text does not hold.
    expect(
      unwrap(await analyzeDocxBlocks(source)).blocks.map(
        ({ text, selfNumbers }) => [text, selfNumbers],
      ),
    ).toEqual([
      ["Annex  - Summary", [6]],
      ["Annex  - Lists", [6]],
    ]);
    const plan = await planFor({
      source,
      template,
      items: (blocks) =>
        blocks.map(({ id, hash }) => ({
          kind: "source" as const,
          blockId: id,
          hash,
          styleId: "HouseBody",
        })),
    });
    const { docx, report } = unwrap(
      await applyTemplatePlan({ source, template, plan }),
    );
    const document = (await readPart(docx, "word/document.xml")) ?? "";

    expect(report.ok).toBe(true);
    // The number the field prints is in no cached text: the field is kept.
    expect(document).toContain(
      '<w:r><w:rPr><w:b/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:rPr><w:b/></w:rPr><w:instrText xml:space="preserve"> AUTONUMLGL  \\* Arabic \\e </w:instrText></w:r><w:r><w:rPr><w:b/></w:rPr><w:fldChar w:fldCharType="end"/></w:r>',
    );
    expect(document).toContain(
      '<w:instrText xml:space="preserve"> AUTONUM </w:instrText>',
    );
    expect(report.dropped.sourceFields).toEqual([]);
  });

  test("a bookmark that ends in a later paragraph marks the one it starts in", async () => {
    const ref = (code: string, result: string): string =>
      `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve">${code}</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${r(result)}<w:r><w:fldChar w:fldCharType="end"/></w:r>`;
    const source = await buildMinimalDocx(
      [
        p(
          `${r("See ")}${ref(" REF whole \\r \\h ", "A")}${ref(" REF over \\r \\h ", "B")}`,
        ),
        // Ends in front of anything of the next paragraph: the heading's.
        p('<w:bookmarkStart w:id="1" w:name="whole"/>'),
        p(`<w:bookmarkEnd w:id="1"/>${r("Schedule of purchasers")}`),
        // Ends behind text of the next paragraph (a heading run in with its
        // text): the paragraph it starts in, not all of it.
        p(`<w:bookmarkStart w:id="2" w:name="over"/>${r("Heading")}`),
        p(`${r("Body")}<w:bookmarkEnd w:id="2"/>`),
      ].join(""),
    );
    const { blocks } = unwrap(await analyzeDocxBlocks(source));

    expect(
      blocks[0]?.crossReferences?.map(
        ({ bookmark, targetBlockId, targetWhole }) => ({
          bookmark,
          targetBlockId,
          targetWhole,
        }),
      ),
    ).toEqual([
      { bookmark: "whole", targetBlockId: "p1", targetWhole: true },
      { bookmark: "over", targetBlockId: "p3", targetWhole: false },
    ]);
  });

  test("a bookmark behind a number a field counts does not hold the whole paragraph", async () => {
    const number =
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> AUTONUM </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>';
    const ref =
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> REF fees \\h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>Fees</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>';
    const source = await buildMinimalDocx(
      [
        p(
          `${number}<w:bookmarkStart w:id="1" w:name="fees"/>${r("Fees")}<w:bookmarkEnd w:id="1"/>`,
        ),
        p(`${r("See ")}${ref}`),
      ].join(""),
    );
    const { blocks } = unwrap(await analyzeDocxBlocks(source));

    // The field's number is content of the paragraph the bookmark leaves out.
    expect(blocks[1]?.crossReferences?.[0]).toMatchObject({
      targetBlockId: "p0",
      targetWhole: false,
    });
  });

  test("keeps the source's own list where the plan asks for it", async () => {
    const template = await buildTemplate();
    const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    // List 7 links a numbering style; the definition with its levels is
    // the one the style names. List 8 counts in it too, from 5 on.
    const numbering = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="3"><w:nsid w:val="11111111"/><w:multiLevelType w:val="multilevel"/><w:styleLink w:val="RowNumbers"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="Missing"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="360" w:hanging="360"/></w:pPr></w:lvl><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="(%2)"/><w:lvlJc w:val="left"/></w:lvl></w:abstractNum><w:abstractNum w:abstractNumId="4"><w:nsid w:val="22222222"/><w:multiLevelType w:val="multilevel"/><w:numStyleLink w:val="RowNumbers"/></w:abstractNum><w:num w:numId="7"><w:abstractNumId w:val="4"/></w:num><w:num w:numId="8"><w:abstractNumId w:val="4"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/></w:lvlOverride></w:num></w:numbering>`;
    const source = await buildMinimalDocx(
      [
        p(r("First"), numPr(7, 0)),
        p(r("Its item"), numPr(7, 1)),
        p(r("Second"), numPr(7, 0)),
        p(r("Fifth"), numPr(8, 0)),
      ].join(""),
      {
        "word/_rels/document.xml.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>`,
        "word/numbering.xml": numbering,
      },
    );
    const sourceLabels = unwrap(await analyzeDocxBlocks(source)).blocks.map(
      ({ label }) => label,
    );
    expect(sourceLabels).toEqual(["1.", "(a)", "2.", "5."]);
    const plan = await planFor({
      source,
      template,
      items: (blocks) =>
        blocks.map(({ id, hash }, index) => ({
          kind: "source" as const,
          blockId: id,
          hash,
          styleId: "HouseBody",
          numbering: {
            numId: "new" as const,
            ilvl: index === 1 ? 1 : 0,
            sourceNumId: index === 3 ? 8 : 7,
          },
        })),
    });
    const { docx, report } = unwrap(
      await applyTemplatePlan({ source, template, plan }),
    );
    const numberingOut = (await readPart(docx, "word/numbering.xml")) ?? "";

    expect(report.ok).toBe(true);
    expect(report.counts.labelChanged).toBe(0);
    expect(report.counts.labelLost).toBe(0);
    expect(
      unwrap(await analyzeDocxBlocks(docx)).blocks.map(({ label }) => label),
    ).toEqual(sourceLabels);
    // One definition, the one with the levels, without its package ties;
    // two instances, the second with its override.
    expect(numberingOut.match(/w:lvlText w:val="%1\."/gu)).toHaveLength(1);
    expect(numberingOut).not.toContain("RowNumbers");
    expect(numberingOut).not.toContain("Missing");
    expect(numberingOut).not.toContain("11111111");
    expect(numberingOut).toContain('<w:startOverride w:val="5"/>');
    // A source list never stands beside a list of the template.
    expect(
      (
        await applyTemplatePlan({
          source,
          template,
          plan: {
            ...plan,
            body: plan.body.map((item) =>
              item.kind === "source"
                ? Object.assign(structuredClone(item), {
                    numbering: { numId: 1, ilvl: 0, sourceNumId: 7 },
                  })
                : item,
            ),
          },
        })
      ).isErr(),
    ).toBe(true);
  });

  test("counts no label lost for a paragraph whose level prints nothing", async () => {
    const template = await buildTemplate();
    const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const source = await buildMinimalDocx(
      p(r("Counted, shown bare"), numPr(5, 0)),
      {
        "word/_rels/document.xml.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>`,
        "word/numbering.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val=""/></w:lvl></w:abstractNum><w:num w:numId="5"><w:abstractNumId w:val="1"/></w:num></w:numbering>`,
      },
    );
    const plan = await planFor({
      source,
      template,
      items: (blocks) =>
        blocks.map(({ id, hash }) => ({
          kind: "source" as const,
          blockId: id,
          hash,
          styleId: "HouseBody",
          numbering: null,
        })),
    });
    const { report } = unwrap(
      await applyTemplatePlan({ source, template, plan }),
    );
    expect(report.counts.labelLost).toBe(0);
  });

  test("keeps a field that counts in a list by name, with the list", async () => {
    const template = await buildTemplate();
    const listnum =
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> LISTNUM  AnnexNumbers </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>';
    const source = await buildMinimalDocx(
      [
        p(`${r("Annex ")}${listnum}${r(" - Prices")}`),
        p(`${r("Annex ")}${listnum}${r(" - Terms")}`),
      ].join(""),
      {
        "word/_rels/document.xml.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>`,
        "word/numbering.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"><w:abstractNum w:abstractNumId="4" w15:restartNumberingAfterBreak="0"><w:nsid w:val="0A534828"/><w:multiLevelType w:val="multilevel"/><w:tmpl w:val="C024B46A"/><w:name w:val="AnnexNumbers"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="Missing"/><w:lvlText w:val="%1"/><w:lvlJc w:val="left"/><w:rPr><w:b/></w:rPr></w:lvl></w:abstractNum><w:num w:numId="7"><w:abstractNumId w:val="4"/></w:num></w:numbering>`,
      },
    );
    const plan = await planFor({
      source,
      template,
      items: (blocks) =>
        blocks.map(({ id, hash }) => ({
          kind: "source" as const,
          blockId: id,
          hash,
          styleId: "HouseBody",
        })),
    });
    const { docx, report } = unwrap(
      await applyTemplatePlan({ source, template, plan }),
    );
    const document = (await readPart(docx, "word/document.xml")) ?? "";
    const numbering = (await readPart(docx, "word/numbering.xml")) ?? "";

    expect(report.ok).toBe(true);
    expect(report.dropped.sourceFields).toEqual([]);
    expect(document.match(/ LISTNUM {2}AnnexNumbers /gu)).toHaveLength(2);
    // The list goes with its fields, without what ties it to its package,
    // and with an instance after the template's.
    expect(numbering).toContain(
      '<w:abstractNum w:abstractNumId="2"><w:multiLevelType w:val="multilevel"/><w:name w:val="AnnexNumbers"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1"/><w:lvlJc w:val="left"/><w:rPr><w:b/></w:rPr></w:lvl></w:abstractNum><w:num w:numId="1">',
    );
    expect(numbering).toContain(
      '<w:num w:numId="3"><w:abstractNumId w:val="2"/></w:num></w:numbering>',
    );
  });

  test("reads a bookmark that ends between cells as ending behind the paragraph before", async () => {
    const ref = (code: string, result: string): string =>
      `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve">${code}</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${r(result)}<w:r><w:fldChar w:fldCharType="end"/></w:r>`;
    // A numbered row cell: the bookmark starts in its empty paragraph and
    // ends in the row, in front of the next cell.
    const source = await buildMinimalDocx(
      [
        `<w:tbl><w:tblGrid><w:gridCol w:w="500"/><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc>${p('<w:bookmarkStart w:id="7" w:name="row1"/>')}</w:tc><w:bookmarkEnd w:id="7"/><w:tc>${p(r("Customer"))}</w:tc></w:tr></w:tbl>`,
        p(`${r("See item ")}${ref(" REF row1 \\w \\h ", "1")}${r(".")}`),
      ].join(""),
    );
    const { blocks } = unwrap(await analyzeDocxBlocks(source));
    const citing = blocks.find(({ text }) => text.startsWith("See item"));

    expect(citing?.crossReferences).toEqual([
      expect.objectContaining({
        bookmark: "row1",
        // The empty paragraph of the first cell.
        targetBlockId: blocks.at(0)?.cells?.at(0)?.blockIds.at(0),
        targetWhole: true,
      }),
    ]);
  });

  test("reports a cross-reference field with the paragraph its bookmark marks", async () => {
    const template = await buildTemplate();
    const number =
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> AUTONUMLGL  \\* Arabic \\e </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>';
    const ref = (code: string, result: string): string =>
      `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve">${code}</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${r(result)}<w:r><w:fldChar w:fldCharType="end"/></w:r>`;
    const source = await buildMinimalDocx(
      [
        p(
          `<w:bookmarkStart w:id="1" w:name="annex1"/>${r("Annex ")}${number}${r(" - Summary")}<w:bookmarkEnd w:id="1"/>`,
        ),
        p(
          `${r("See ")}${ref(" REF annex1 \\h  \\* MERGEFORMAT ", "Annex 1 - Summary")}${r(".")}`,
        ),
        p(
          `${r("Also ")}${ref(" REF part \\h ", "Summary")}${r(" and ")}${ref(" REF gone \\h ", "Annex 9")}`,
        ),
        p(
          `${r("Annex ")}<w:bookmarkStart w:id="2" w:name="part"/>${r("Summary")}<w:bookmarkEnd w:id="2"/>`,
        ),
      ].join(""),
    );
    const { blocks } = unwrap(await analyzeDocxBlocks(source));

    expect(blocks.map(({ crossReferences }) => crossReferences)).toEqual([
      undefined,
      [
        {
          start: 4,
          end: 21,
          instruction: " REF annex1 \\h  \\* MERGEFORMAT ",
          bookmark: "annex1",
          targetBlockId: "p0",
          targetWhole: true,
        },
      ],
      [
        {
          start: 5,
          end: 12,
          instruction: " REF part \\h ",
          bookmark: "part",
          targetBlockId: "p3",
          targetWhole: false,
        },
        {
          start: 17,
          end: 24,
          instruction: " REF gone \\h ",
          bookmark: "gone",
          targetBlockId: null,
          targetWhole: false,
        },
      ],
      undefined,
    ]);

    // Kept as a field, the reference shows the heading's text with the
    // number the heading's field counts.
    const plan = await planFor({
      source,
      template,
      extra: {
        references: [
          {
            blockId: "p1",
            start: 4,
            end: 21,
            targetBlockId: "p0",
            // An update keeps the result's look, as it did in the source.
            instruction: "\\h \\* MERGEFORMAT",
            expectedText: "Annex 1 - Summary",
          },
        ],
      },
    });
    const { docx, report } = unwrap(
      await applyTemplatePlan({ source, template, plan }),
    );
    const document = (await readPart(docx, "word/document.xml")) ?? "";
    const [result] = report.references.results;

    expect(report.ok).toBe(true);
    expect(result).toMatchObject({
      status: "ok",
      cachedText: "Annex 1 - Summary",
      replacedText: "Annex 1 - Summary",
      // The number is counted when the document is laid out.
      predictedText: null,
    });
    expect(report.references.predictionMismatch).toBe(0);
    // The field the plan writes over the source's stands for it; the two
    // the plan leaves are written as their text.
    expect(report.dropped.sourceFields).toEqual([{ kind: "REF", count: 2 }]);
    expect(document).toContain(
      `<w:instrText xml:space="preserve"> REF ${result?.bookmark ?? ""} \\h \\* MERGEFORMAT </w:instrText>`,
    );
    expect(document).toContain(" AUTONUMLGL  \\* Arabic \\e ");
  });

  test("a page reference is reported and kept as a field to the plan's bookmark on its target", async () => {
    const template = await buildTemplate();
    const field = (code: string, result: string): string =>
      `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve">${code}</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${r(result)}<w:r><w:fldChar w:fldCharType="end"/></w:r>`;
    const source = await buildMinimalDocx(
      [
        p(
          `<w:bookmarkStart w:id="1" w:name="a19"/>${r("Dispute resolution")}<w:bookmarkEnd w:id="1"/>`,
        ),
        // A picture of words alone: the field shows them and links to the
        // heading, as a drafter writes "clause " in front of its number.
        p(
          `${r("As set out in ")}${field("PAGEREF a19\\# \"'clause '\"  \\h", "clause ")}${r("19.")}`,
        ),
      ].join(""),
    );
    const { blocks } = unwrap(await analyzeDocxBlocks(source));
    expect(blocks[1]?.crossReferences).toMatchObject([
      { start: 14, end: 21, bookmark: "a19", targetBlockId: "p0" },
    ]);
    const plan = await planFor({
      source,
      template,
      extra: {
        references: [
          {
            blockId: "p1",
            start: 14,
            end: 21,
            targetBlockId: "p0",
            instruction: `\\# "'clause '" \\h`,
            expectedText: "clause ",
            field: "PAGEREF",
          },
        ],
      },
    });
    const { docx, report } = unwrap(
      await applyTemplatePlan({ source, template, plan }),
    );
    const document = (await readPart(docx, "word/document.xml")) ?? "";
    const [result] = report.references.results;
    expect(report.ok).toBe(true);
    expect(result).toMatchObject({ status: "ok", predictedText: null });
    expect(document).toContain(
      `<w:instrText xml:space="preserve"> PAGEREF ${result?.bookmark ?? ""} \\# &quot;&apos;clause &apos;&quot; \\h </w:instrText>`,
    );
    // A switch a page reference cannot carry refuses the plan.
    const odd = await applyTemplatePlan({
      source,
      template,
      plan: {
        ...plan,
        references: (plan.references ?? []).map((reference) =>
          Object.assign({}, reference, { instruction: "\\n" }),
        ),
      },
    });
    expect(odd.isErr()).toBe(true);
  });

  test("a link to a bookmark of the source leads to the output's target when the plan keeps it", async () => {
    const template = await buildTemplate();
    const field = (code: string, result: string): string =>
      `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve">${code}</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${r(result)}<w:r><w:fldChar w:fldCharType="end"/></w:r>`;
    const source = await buildMinimalDocx(
      [
        p(`<w:hyperlink w:anchor="_Toc1">${r("Definitions")}</w:hyperlink>`),
        p(`${r("See ")}${field(' HYPERLINK \\l "_Toc1" ', "the definitions")}`),
        p(
          `<w:bookmarkStart w:id="1" w:name="_Toc1"/>${r("Definitions")}<w:bookmarkEnd w:id="1"/>`,
        ),
        p(`<w:hyperlink w:anchor="nowhere">${r("Lost")}</w:hyperlink>`),
      ].join(""),
    );
    const kept = unwrap(
      await applyFor({ source, template, extra: { internalLinks: "keep" } }),
    );
    const document = (await readPart(kept.docx, "word/document.xml")) ?? "";
    const bookmark =
      /<w:bookmarkStart w:id="\d+" w:name="([^"]+)"\/><w:r>(?:<w:rPr>.*?<\/w:rPr>)?<w:t xml:space="preserve">Definitions/u.exec(
        document,
      )?.[1];
    expect(bookmark).toBeDefined();
    expect(
      document.match(
        new RegExp(`<w:hyperlink w:anchor="${bookmark ?? ""}"`, "gu"),
      )?.length,
    ).toBe(2);
    expect(kept.report.ok).toBe(true);
    // The link to a bookmark the source lacks led nowhere there: it stays
    // text, and no link is lost.
    expect(kept.report.dropped.internalHyperlinks).toBe(0);
    const plain = unwrap(await applyFor({ source, template }));
    expect(plain.report.dropped.internalHyperlinks).toBe(2);
    expect(
      (await readPart(plain.docx, "word/document.xml")) ?? "",
    ).not.toContain("w:anchor");
  });

  test("a contents entry field and a macro button stay fields where their paragraph is placed", async () => {
    const template = await buildTemplate();
    const code = (instruction: string): string =>
      `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve">${instruction}</w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;
    const source = await buildMinimalDocx(
      [
        p(`${code(' TC "12. Service credits" \\l 1 ')}${r("Service credits")}`),
        p(
          `${code(" MACROBUTTON optional [ ")}${r("The Supplier shall notify.]")}`,
        ),
      ].join(""),
    );
    const { docx, report } = unwrap(await applyFor({ source, template }));
    const document = (await readPart(docx, "word/document.xml")) ?? "";
    expect(report.ok).toBe(true);
    expect(document).toContain(
      '<w:instrText xml:space="preserve"> TC &quot;12. Service credits&quot; \\l 1 </w:instrText>',
    );
    expect(document).toContain(
      '<w:instrText xml:space="preserve"> MACROBUTTON optional [ </w:instrText>',
    );
    expect(report.dropped.sourceFields).toEqual([]);
  });

  test("a colour that marks part of the text counts the text of the notes", async () => {
    const template = await buildTemplate();
    const separators =
      '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>';
    const source = await buildSourceWith({
      body: p(
        `${r("[Insert date]", '<w:color w:val="FF0000"/>')}<w:r><w:footnoteReference w:id="2"/></w:r>`,
      ),
      footnotes: `${separators}<w:footnote w:id="2">${p(r(" The date on which the last party signs this agreement."))}</w:footnote>`,
    });
    const kept = unwrap(
      await applyFor({ source, template, extra: { textColors: "keep" } }),
    );
    expect((await readPart(kept.docx, "word/document.xml")) ?? "").toContain(
      '<w:color w:val="FF0000"/>',
    );
  });

  test("a colour that marks part of the text is kept when the plan keeps colours; the source's own look is not", async () => {
    const template = await buildTemplate();
    const source = await buildMinimalDocx(
      [
        p(
          `${r("The Supplier shall deliver the goods on time and in full. ", '<w:color w:val="1F3864"/>')}${r("[Insert date]", '<w:color w:val="FF0000"/><w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/>')}`,
        ),
        p(
          r(
            "Each party bears its own costs of this agreement.",
            '<w:color w:val="1F3864"/>',
          ),
        ),
      ].join(""),
    );
    const kept = unwrap(
      await applyFor({ source, template, extra: { textColors: "keep" } }),
    );
    const document = (await readPart(kept.docx, "word/document.xml")) ?? "";
    expect(kept.report.ok).toBe(true);
    expect(document).toContain(
      '<w:color w:val="FF0000"/><w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/></w:rPr><w:t xml:space="preserve">[Insert date]',
    );
    // The blue that covers most of the text is the source's look.
    expect(document).not.toContain("1F3864");
    const plain = unwrap(await applyFor({ source, template }));
    expect(
      (await readPart(plain.docx, "word/document.xml")) ?? "",
    ).not.toContain("w:color");
  });

  test("a marking colour a paragraph style gives is kept, and run boundaries do not change what marks", async () => {
    const template = await buildTemplate();
    const styled = await buildSourceWith({
      body: [
        p(
          r("Note to draft: check the dates."),
          '<w:pStyle w:val="DraftNote"/>',
        ),
        p(r("The Supplier shall deliver the goods on time and in full.")),
        p(
          r(
            "Each party bears its own costs of this agreement and of its advisers.",
          ),
        ),
      ].join(""),
      styles:
        '<w:style w:type="paragraph" w:styleId="DraftNote"><w:name w:val="Draft Note"/><w:rPr><w:color w:val="C00000"/></w:rPr></w:style>',
    });
    const kept = unwrap(
      await applyFor({
        source: styled,
        template,
        extra: { textColors: "keep" },
      }),
    );
    expect((await readPart(kept.docx, "word/document.xml")) ?? "").toMatch(
      /<w:color w:val="C00000"\/><\/w:rPr><w:t xml:space="preserve">Note to draft/u,
    );
    const colours = async (runs: string): Promise<boolean> => {
      const source = await buildMinimalDocx(p(`${runs}${r("ccc")}`));
      const applied = unwrap(
        await applyFor({ source, template, extra: { textColors: "keep" } }),
      );
      return (
        (await readPart(applied.docx, "word/document.xml")) ?? ""
      ).includes("FF0000");
    };
    const red = '<w:color w:val="FF0000"/>';
    expect(await colours(r("a b", red))).toBe(
      await colours(`${r("a ", red)}${r("b", red)}`),
    );
  });

  test("a content control among runs is kept around its content, without the parts it names", async () => {
    const template = await buildTemplate();
    const control = (properties: string, text: string): string =>
      `<w:sdt><w:sdtPr>${properties}</w:sdtPr><w:sdtContent>${r(text)}</w:sdtContent></w:sdt>`;
    const source = await buildMinimalDocx(
      [
        p(
          `${r("Signed on ")}${control('<w:id w:val="7"/><w:placeholder><w:docPart w:val="ABC"/></w:placeholder><w:showingPlcHdr/><w:dataBinding w:xpath="/a" w:storeItemID="{1}"/><w:date><w:dateFormat w:val="d MMMM yyyy"/><w:lid w:val="en-GB"/></w:date>', "Click to enter a date.")}${r(" by the parties.")}`,
        ),
        p(
          `${r("Name: ")}${control('<w:rPr><w:rStyle w:val="instructions"/></w:rPr><w:alias w:val="Name"/><w:text/>', "[Name]")}`,
        ),
      ].join(""),
    );
    const { docx, report } = unwrap(await applyFor({ source, template }));
    const document = (await readPart(docx, "word/document.xml")) ?? "";
    expect(report.ok).toBe(true);
    expect(document).toContain(
      '<w:sdt><w:sdtPr><w:id w:val="7"/><w:showingPlcHdr/><w:date><w:dateFormat w:val="d MMMM yyyy"/><w:lid w:val="en-GB"/></w:date></w:sdtPr><w:sdtContent>',
    );
    expect(document).not.toContain("dataBinding");
    expect(document).not.toContain("docPart");
    // Nor the look of a style the template does not define.
    expect(document).not.toContain("instructions");
    expect(document).toMatch(
      /<w:sdt><w:sdtPr><w:alias w:val="Name"\/><w:text\/><\/w:sdtPr><w:sdtContent><w:r>(?:<w:rPr>.*?<\/w:rPr>)?<w:t xml:space="preserve">\[Name\]<\/w:t><\/w:r><\/w:sdtContent><\/w:sdt>/u,
    );
  });

  test("a paragraph whose mark is hidden keeps it hidden, so it runs on into the next", async () => {
    const template = await buildTemplate();
    const source = await buildMinimalDocx(
      [
        p(r("Die Vergütung ist fällig. "), "<w:rPr><w:vanish/></w:rPr>"),
        p(r("Der Auftraggeber zahlt binnen 14 Tagen.")),
      ].join(""),
    );
    const { docx, report } = unwrap(await applyFor({ source, template }));
    const document = (await readPart(docx, "word/document.xml")) ?? "";
    expect(report.ok).toBe(true);
    expect(document).toMatch(
      /<w:rPr><w:vanish\/><\/w:rPr><\/w:pPr><w:r>(?:<w:rPr>.*?<\/w:rPr>)?<w:t xml:space="preserve">Die Vergütung/u,
    );
    expect(document.match(/<w:vanish\/>/gu)?.length).toBe(1);
  });

  test("a content control around paragraphs stays around them where the output writes them all", async () => {
    const template = await buildTemplate();
    const source = await buildMinimalDocx(
      [
        p(r("The parties agree as follows.")),
        `<w:sdt><w:sdtPr><w:alias w:val="Schedule"/><w:dataBinding w:xpath="/a" w:storeItemID="{1}"/></w:sdtPr><w:sdtContent>${p(r("Item one of the schedule."))}${p(r("Item two of the schedule."))}</w:sdtContent></w:sdt>`,
        p(r("Each party signs below.")),
      ].join(""),
    );
    const { docx, report } = unwrap(await applyFor({ source, template }));
    const document = (await readPart(docx, "word/document.xml")) ?? "";
    expect(report.ok).toBe(true);
    expect(document).toMatch(
      /<w:sdt><w:sdtPr><w:alias w:val="Schedule"\/><\/w:sdtPr><w:sdtContent><w:p>.*?Item one of the schedule\..*?<\/w:p><w:p>.*?Item two of the schedule\..*?<\/w:p><\/w:sdtContent><\/w:sdt>/u,
    );
    expect(document).not.toContain("dataBinding");
  });

  test("a content control around paragraphs and an inner control stays around them both", async () => {
    const template = await buildTemplate();
    const source = await buildMinimalDocx(
      [
        p(r("The parties agree as follows.")),
        `<w:sdt><w:sdtPr><w:alias w:val="Schedule"/></w:sdtPr><w:sdtContent>${p(r("Item one of the schedule."))}<w:sdt><w:sdtPr><w:alias w:val="Item"/></w:sdtPr><w:sdtContent>${p(r("Item two of the schedule."))}</w:sdtContent></w:sdt></w:sdtContent></w:sdt>`,
        p(r("Each party signs below.")),
      ].join(""),
    );
    const { docx, report } = unwrap(await applyFor({ source, template }));
    const document = (await readPart(docx, "word/document.xml")) ?? "";
    expect(report.ok).toBe(true);
    expect(document).toMatch(
      /<w:sdt><w:sdtPr><w:alias w:val="Schedule"\/><\/w:sdtPr><w:sdtContent><w:p>.*?Item one of the schedule\..*?<\/w:p><w:sdt><w:sdtPr><w:alias w:val="Item"\/><\/w:sdtPr><w:sdtContent><w:p>.*?Item two of the schedule\..*?<\/w:p><\/w:sdtContent><\/w:sdt><\/w:sdtContent><\/w:sdt>/u,
    );
  });

  test("a paragraph cites its own number through a field", async () => {
    const template = await buildTemplate();
    const source = await buildMinimalDocx(
      p(r("Within this Section 1 the fee applies.")),
    );
    const self = (instruction: string, expectedText: string) =>
      planFor({
        source,
        template,
        items: ([first]) =>
          first === undefined
            ? []
            : [
                {
                  kind: "source",
                  blockId: first.id,
                  hash: first.hash,
                  styleId: "HouseBody",
                  numbering: { numId: 1, ilvl: 0 },
                },
              ],
        extra: {
          references: [
            {
              blockId: "p0",
              start: 20,
              end: 21,
              targetBlockId: "p0",
              instruction,
              expectedText,
            },
          ],
        },
      });
    const { docx, report } = unwrap(
      await applyTemplatePlan({
        source,
        template,
        plan: await self("\\n \\t \\h", "1"),
      }),
    );
    const document = (await readPart(docx, "word/document.xml")) ?? "";
    const [result] = report.references.results;

    // The paragraph's bookmark holds the field that shows its number.
    expect(report.ok).toBe(true);
    expect(result).toMatchObject({
      status: "ok",
      cachedText: "1",
      predictedText: "1",
    });
    expect(report.references.predictionMismatch).toBe(0);
    expect(document).toMatch(
      new RegExp(
        `<w:bookmarkStart w:id="\\d+" w:name="${result?.bookmark ?? ""}"/>.*REF ${result?.bookmark ?? ""} \\\\n \\\\t \\\\h .*<w:bookmarkEnd`,
        "u",
      ),
    );
    // A field that shows the paragraph's text would show itself.
    const looped = await applyTemplatePlan({
      source,
      template,
      plan: await self("\\h", "1"),
    });
    expect(looped.isErr() ? looped.error.code : null).toBe(
      "reference-span-invalid",
    );
  });

  test("strips a typed number whose parts stand apart", async () => {
    const template = await buildTemplate();
    const source = await buildMinimalDocx(
      p(`${r("1. 1.")}<w:r><w:tab/></w:r>${r("The fee is due.")}`),
    );
    const plan = await planFor({
      source,
      template,
      items: ([first]) =>
        first === undefined
          ? []
          : [
              {
                kind: "source",
                blockId: first.id,
                hash: first.hash,
                styleId: "HouseBody",
                stripLabel: "1. 1.",
                numbering: { numId: 1, ilvl: 1 },
              },
            ],
    });
    const { docx, report } = unwrap(
      await applyTemplatePlan({ source, template, plan }),
    );
    const document = (await readPart(docx, "word/document.xml")) ?? "";

    expect(report.ok).toBe(true);
    expect(document).toContain(
      '<w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t xml:space="preserve">The fee is due.</w:t></w:r></w:p>',
    );
    expect(report.blocks.at(0)).toMatchObject({ sourceLabel: "1. 1." });
  });

  test("carries symbols, hyphens and breaks, and strips a label after leading space", async () => {
    const template = await buildTemplate();
    const source = await buildMinimalDocx(
      [
        p(
          `${r("  (a)")}<w:r><w:tab/></w:r>${r("non")}<w:r><w:noBreakHyphen/></w:r>${r("breaking ")}<w:r><w:sym w:font="Wingdings" w:char="F0FC"/></w:r><w:r><w:br w:type="column"/></w:r>${r("after")}`,
        ),
      ].join(""),
    );
    const plan = await planFor({
      source,
      template,
      items: ([first]) =>
        first === undefined
          ? []
          : [
              {
                kind: "source",
                blockId: first.id,
                hash: first.hash,
                styleId: "HouseBody",
                stripLabel: "(a)",
              },
            ],
    });
    const { docx, report } = unwrap(
      await applyTemplatePlan({ source, template, plan }),
    );
    const document = (await readPart(docx, "word/document.xml")) ?? "";

    expect(report.blocks.at(0)).toMatchObject({
      text: "identical",
      strippedPrefix: "  (a)\t",
      sourceLabel: "(a)",
    });
    expect(unwrap(await analyzeDocxBlocks(docx)).blocks.at(0)?.text).toBe(
      "non‑breaking after",
    );
    expect(document).toContain("<w:noBreakHyphen/>");
    expect(document).toContain('<w:sym w:font="Wingdings" w:char="F0FC"/>');
    expect(document).toContain('<w:br w:type="column"/>');

    // The same text with the break after it is another reading.
    const moved = document.replace(
      /(<w:r><w:br w:type="column"\/><\/w:r>)(<w:r>.*?<\/w:r>)/u,
      "$2$1",
    );
    expect(moved).not.toBe(document);
    const movedReport = await verifyFor({
      source,
      template,
      plan,
      output: await replacePart(docx, "word/document.xml", moved),
    });
    expect(movedReport.issues.map(({ code }) => code)).toEqual([
      "text-mismatch",
    ]);
  });

  test("drops a bookmark half whose other half sits in a dropped template block", async () => {
    const base = await buildTemplate();
    const xml = (await readPart(base, "word/document.xml")) ?? "";
    // The bookmark starts in the title and ends in the party paragraph.
    const template = await replacePart(
      base,
      "word/document.xml",
      xml
        .replace(
          '<w:t xml:space="preserve">Agreement</w:t>',
          '<w:t xml:space="preserve">Agreement</w:t></w:r><w:bookmarkStart w:id="40" w:name="Spanning"/><w:r>',
        )
        .replace(
          '<w:t xml:space="preserve">[party name]</w:t></w:r>',
          '<w:t xml:space="preserve">[party name]</w:t></w:r><w:bookmarkEnd w:id="40"/>',
        ),
    );
    const source = await buildMinimalDocx(p(r("Body text.")));
    const plan = await planFor({
      source,
      template,
      items: ([first]) => [
        { kind: "template", templateBlockIds: [TEMPLATE_IDS.title] },
        ...(first === undefined
          ? []
          : [
              {
                kind: "source" as const,
                blockId: first.id,
                hash: first.hash,
              },
            ]),
      ],
    });
    const { docx, report } = unwrap(
      await applyTemplatePlan({ source, template, plan }),
    );
    const document = (await readPart(docx, "word/document.xml")) ?? "";

    expect(report.ok).toBe(true);
    expect(document).toContain("Agreement");
    expect(document).not.toContain("Spanning");
    expect(document).not.toContain("bookmarkEnd");
  });

  test("patches a self-closed level and starts a list where the part holds none", async () => {
    const base = await buildTemplate();
    const numbering = (await readPart(base, "word/numbering.xml")) ?? "";
    const template = await replacePart(
      base,
      "word/numbering.xml",
      numbering.replace(
        /<w:lvl w:ilvl="4">.*?<\/w:lvl>/u,
        '<w:lvl w:ilvl="4"/>',
      ),
    );
    const source = await buildMinimalDocx(`${p(r("first"))}${p(r("second"))}`);
    const plan = await planFor({
      source,
      template,
      items: (ids) =>
        ids.map(({ id, hash }) => ({
          kind: "source",
          blockId: id,
          hash,
          styleId: "HouseBody",
          numbering: { numId: 1, ilvl: 4, restartFrom: 2, group: "romans" },
        })),
      extra: {
        numberingPatches: [
          {
            abstractNumId: 0,
            ilvl: 4,
            numFmt: "lowerRoman",
            lvlText: "(%5)",
            start: 1,
            suff: "space",
          },
        ],
      },
    });
    const { docx } = unwrap(
      await applyTemplatePlan({ source, template, plan }),
    );
    const patched = (await readPart(docx, "word/numbering.xml")) ?? "";

    expect(patched).toContain(
      '<w:lvl w:ilvl="4"><w:start w:val="1"/><w:numFmt w:val="lowerRoman"/><w:suff w:val="space"/><w:lvlText w:val="(%5)"/></w:lvl>',
    );
    expect(
      unwrap(await analyzeDocxBlocks(docx)).blocks.map(({ label }) => label),
    ).toEqual(["(ii)", "(iii)"]);

    // A list that starts elsewhere, or a paragraph at another level, is not
    // the numbering the plan asked for.
    const restarted = patched.replace(
      '<w:startOverride w:val="2"/>',
      '<w:startOverride w:val="1"/>',
    );
    expect(restarted).not.toBe(patched);
    const restartedReport = await verifyFor({
      source,
      template,
      plan,
      output: await replacePart(docx, "word/numbering.xml", restarted),
    });
    expect(restartedReport.issues).toEqual([
      expect.objectContaining({
        code: "numbering-mismatch",
        part: "word/numbering.xml",
      }),
    ]);

    const document = (await readPart(docx, "word/document.xml")) ?? "";
    const relevelled = document.replace(
      '<w:ilvl w:val="4"/>',
      '<w:ilvl w:val="3"/>',
    );
    expect(relevelled).not.toBe(document);
    const relevelledReport = await verifyFor({
      source,
      template,
      plan,
      output: await replacePart(docx, "word/document.xml", relevelled),
    });
    expect(relevelledReport.issues).toEqual([
      expect.objectContaining({ code: "numbering-mismatch", blockId: "p0" }),
    ]);
  });

  test("patches the indent of a level and keeps what the patch leaves out", async () => {
    const base = await buildTemplate();
    const numbering = (await readPart(base, "word/numbering.xml")) ?? "";
    // Level 0 states an indent; level 1 holds paragraph properties without
    // one; level 4 is written self-closed.
    const template = await replacePart(
      base,
      "word/numbering.xml",
      numbering
        .replace(
          '<w:pPr><w:ind w:left="567" w:hanging="567"/></w:pPr>',
          '<w:pPr><w:ind w:left="567" w:right="2mm" w:hanging="567"/><w:jc w:val="center"/></w:pPr>',
        )
        .replace(
          '<w:pPr><w:ind w:left="1134" w:hanging="567"/></w:pPr>',
          '<w:pPr><w:keepNext/><w:jc w:val="left"/></w:pPr>',
        )
        // A level whose properties are written self-closed, with a namespace
        // declared on the tag itself, and one that states its indent in inches.
        .replace(
          '<w:pPr><w:ind w:left="2268" w:hanging="567"/></w:pPr>',
          '<w:pPr xmlns:x="urn:example:local"/>',
        )
        .replace(/<w:lvl w:ilvl="4">.*?<\/w:lvl>/u, '<w:lvl w:ilvl="4"/>'),
    );
    expect(template).not.toBe(base);
    const inches = await replacePart(
      base,
      "word/numbering.xml",
      numbering
        .replace(
          '<w:ind w:left="567" w:hanging="567"/>',
          '<w:ind w:left="1in" w:hanging="0.5in"/>',
        )
        // Ids may be spelled with leading zeros.
        .replace('<w:lvl w:ilvl="0">', '<w:lvl w:ilvl="00">'),
    );
    expect(inches).not.toBe(base);
    expect(
      unwrap(await inspectDocxTemplate(inches))
        .numbering.abstractNums.at(0)
        ?.levels.at(0)?.indent,
    ).toEqual({ left: 1440, hanging: 720, firstLine: null });
    const source = await buildMinimalDocx(`${p(r("first"))}${p(r("second"))}`);
    const plan = await planFor({
      source,
      template,
      items: (ids) =>
        ids.map(({ id, hash }) => ({
          kind: "source",
          blockId: id,
          hash,
          styleId: "HouseBody",
          numbering: { numId: 1, ilvl: 0 },
        })),
      extra: {
        numberingPatches: [
          { abstractNumId: 0, ilvl: 0, indent: { left: 0, hanging: 0 } },
          { abstractNumId: 0, ilvl: 1, indent: { left: 284 } },
          { abstractNumId: 0, ilvl: 2, indent: { firstLine: 142 } },
          { abstractNumId: 0, ilvl: 3, indent: { hanging: 284 } },
          { abstractNumId: 0, ilvl: 4, lvlText: "%5.", indent: { left: 0 } },
        ],
      },
    });
    const { docx, report } = unwrap(
      await applyTemplatePlan({ source, template, plan }),
    );
    const patched = (await readPart(docx, "word/numbering.xml")) ?? "";

    expect(report.ok).toBe(true);
    // The right indent stays as written, in its own unit, and so does the
    // alignment; left and hanging are the plan's.
    expect(patched).toContain(
      '<w:pPr><w:ind w:left="0" w:right="2mm" w:hanging="0"/><w:jc w:val="center"/></w:pPr>',
    );
    // An indent is written where the schema orders it: before the alignment.
    expect(patched).toContain(
      '<w:pPr><w:keepNext/><w:ind w:left="284"/><w:jc w:val="left"/></w:pPr>',
    );
    // A first-line indent replaces the hanging one; the left indent stays.
    expect(patched).toContain(
      '<w:pPr><w:ind w:left="1701" w:firstLine="142"/></w:pPr>',
    );
    // A self-closed `w:pPr` keeps its start tag as written.
    expect(patched).toContain(
      '<w:pPr xmlns:x="urn:example:local"><w:ind w:hanging="284"/></w:pPr>',
    );
    expect(patched).toContain(
      '<w:lvl w:ilvl="4"><w:lvlText w:val="%5."/><w:pPr><w:ind w:left="0"/></w:pPr></w:lvl>',
    );
    expect(
      unwrap(await inspectDocxTemplate(docx))
        .numbering.abstractNums.at(0)
        ?.levels.map(({ indent }) => indent),
    ).toEqual([
      { left: 0, hanging: 0, firstLine: null },
      { left: 284, hanging: null, firstLine: null },
      { left: 1701, hanging: null, firstLine: 142 },
      { left: null, hanging: 284, firstLine: null },
      { left: 0, hanging: null, firstLine: null },
    ]);

    // An output whose level lost the planned indent is not what was planned.
    const reverted = patched.replace(
      '<w:ind w:left="0" w:right="2mm" w:hanging="0"/>',
      '<w:ind w:left="567" w:right="2mm" w:hanging="567"/>',
    );
    expect(reverted).not.toBe(patched);
    const revertedReport = await verifyFor({
      source,
      template,
      plan,
      output: await replacePart(docx, "word/numbering.xml", reverted),
    });
    expect(
      revertedReport.issues.map(({ code, part }) => [code, part]),
    ).toContainEqual(["numbering-mismatch", "word/numbering.xml"]);
  });

  test("sets a level's label in a face and keeps its other run properties", async () => {
    const base = await buildTemplate();
    const numbering = (await readPart(base, "word/numbering.xml")) ?? "";
    const indentOf = (ilvl: number): string =>
      `<w:pPr><w:ind w:left="${(ilvl + 1) * 567}" w:hanging="567"/></w:pPr>`;
    // Level 0 states run properties without a face, level 1 names a theme
    // face, level 2 writes its run properties self-closed, level 3 has none.
    const template = await replacePart(
      base,
      "word/numbering.xml",
      numbering
        .replace(
          indentOf(0),
          `${indentOf(0)}<w:rPr><w:rStyle w:val="Mark"/><w:b/></w:rPr>`,
        )
        .replace(
          indentOf(1),
          `${indentOf(1)}<w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/><w:sz w:val="18"/></w:rPr>`,
        )
        .replace(indentOf(2), `${indentOf(2)}<w:rPr/>`),
    );
    expect(template).not.toBe(base);
    const source = await buildMinimalDocx(`${p(r("first"))}${p(r("second"))}`);
    const plan = await planFor({
      source,
      template,
      items: (ids) =>
        ids.map(({ id, hash }) => ({
          kind: "source",
          blockId: id,
          hash,
          styleId: "HouseBody",
          numbering: { numId: 1, ilvl: 3 },
        })),
      extra: {
        numberingPatches: [
          { abstractNumId: 0, ilvl: 0, font: "Courier New" },
          { abstractNumId: 0, ilvl: 1, font: "Courier New" },
          { abstractNumId: 0, ilvl: 2, font: "Courier New" },
          {
            abstractNumId: 0,
            ilvl: 3,
            numFmt: "bullet",
            lvlText: "o",
            font: "Courier New",
          },
        ],
      },
    });
    const { docx, report } = unwrap(
      await applyTemplatePlan({ source, template, plan }),
    );
    const patched = (await readPart(docx, "word/numbering.xml")) ?? "";
    const face =
      '<w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:eastAsia="Courier New" w:cs="Courier New" w:hint="default"/>';

    expect(report.ok).toBe(true);
    // The face follows the character style; the weight stays.
    expect(patched).toContain(
      `<w:rPr><w:rStyle w:val="Mark"/>${face}<w:b/></w:rPr>`,
    );
    // A theme face would override the named one: it is replaced, the size stays.
    expect(patched).toContain(`<w:rPr>${face}<w:sz w:val="18"/></w:rPr>`);
    expect(patched).toContain(`${indentOf(2)}<w:rPr>${face}</w:rPr>`);
    expect(patched).toContain(
      `<w:numFmt w:val="bullet"/><w:pStyle w:val="HouseLetter"/><w:lvlText w:val="o"/><w:lvlJc w:val="left"/>${indentOf(3)}<w:rPr>${face}</w:rPr></w:lvl>`,
    );
    // Set in that face the mark reads as the letter it is, and says so.
    expect(
      unwrap(await analyzeDocxBlocks(docx)).blocks.map(
        ({ label, numbering: listed }) => [label, listed?.bulletFont],
      ),
    ).toEqual([
      ["o", "Courier New"],
      ["o", "Courier New"],
    ]);

    // A face for one script only is not the planned face.
    const mixed = patched.replace(
      `${indentOf(2)}<w:rPr>${face}</w:rPr>`,
      `${indentOf(2)}<w:rPr>${face.replace('w:eastAsia="Courier New"', 'w:eastAsia="Serif"')}</w:rPr>`,
    );
    expect(mixed).not.toBe(patched);
    expect(
      (
        await verifyFor({
          source,
          template,
          plan,
          output: await replacePart(docx, "word/numbering.xml", mixed),
        })
      ).issues.map(({ code, part }) => [code, part]),
    ).toContainEqual(["numbering-mismatch", "word/numbering.xml"]);

    // An output whose level lost the face is not what was planned.
    const reverted = patched.replace(
      `<w:rPr>${face}</w:rPr></w:lvl>`,
      "</w:lvl>",
    );
    expect(reverted).not.toBe(patched);
    const revertedReport = await verifyFor({
      source,
      template,
      plan,
      output: await replacePart(docx, "word/numbering.xml", reverted),
    });
    expect(
      revertedReport.issues.map(({ code, part }) => [code, part]),
    ).toContainEqual(["numbering-mismatch", "word/numbering.xml"]);
  });

  test("a line break behind a stripped label goes with it when the plan says so", async () => {
    const template = await buildTemplate();
    const source = await buildMinimalDocx(
      p(`${r("Article I")}<w:r><w:br/></w:r>${r(" First things")}`),
    );
    const planWith = async (stripLabelBreak: boolean) =>
      await planFor({
        source,
        template,
        items: (ids) =>
          ids.map(({ id, hash }) => ({
            kind: "source",
            blockId: id,
            hash,
            styleId: "HouseArticle",
            numbering: { numId: 1, ilvl: 0 },
            stripLabel: "Article I",
            ...(stripLabelBreak ? { stripLabelBreak } : {}),
          })),
      });
    const textOf = async (docx: Uint8Array) =>
      unwrap(await analyzeDocxBlocks(docx)).blocks.map(({ text }) => text);

    // Left out, the break stays the paragraph's first character.
    const kept = unwrap(
      await applyTemplatePlan({
        source,
        template,
        plan: await planWith(false),
      }),
    );
    expect(kept.report.ok).toBe(true);
    expect(await textOf(kept.docx)).toEqual(["\n First things"]);

    // Asked for, the break and the blank around it go with the label.
    const plan = await planWith(true);
    const joined = unwrap(await applyTemplatePlan({ source, template, plan }));
    expect(joined.report.ok).toBe(true);
    expect(await textOf(joined.docx)).toEqual(["First things"]);
    expect(joined.report.blocks.at(0)?.strippedPrefix).toBe("Article I\n ");

    // The two outputs are not each other's: verification tells them apart.
    const mixed = await verifyFor({
      source,
      template,
      plan,
      output: kept.docx,
    });
    expect(mixed.ok).toBe(false);
  });
});
