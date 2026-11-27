import { describe, expect, test } from "bun:test";

import { applyTemplatePlan } from "../apply";
import { analyzeDocxBlocks } from "../blocks";
import { inspectDocxTemplate } from "../inspect";
import { innerRangeOf, readRawPart } from "../xml";
import {
  buildMinimalDocx,
  buildTemplate,
  p,
  r,
  readPart,
  replacePart,
  TEMPLATE_IDS,
} from "./fixtures";
import { planFor, unwrap, verifyFor } from "./planning";

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
    expect(document).toContain(
      `<w:instrText xml:space="preserve"> REF ${result?.bookmark ?? ""} \\h \\* MERGEFORMAT </w:instrText>`,
    );
    expect(document).toContain(" AUTONUMLGL  \\* Arabic \\e ");
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
