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
      numbering.replace(
        '<w:ind w:left="567" w:hanging="567"/>',
        '<w:ind w:left="1in" w:hanging="0.5in"/>',
      ),
    );
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
});
