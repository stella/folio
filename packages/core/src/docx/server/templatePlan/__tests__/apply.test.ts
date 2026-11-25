import { Result } from "better-result";
import { beforeAll, describe, expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { applyTemplatePlan } from "../apply";
import { analyzeDocxBlocks } from "../blocks";
import { inspectDocxTemplate } from "../inspect";
import type {
  DocxBlock,
  DocxBlockManifest,
  DocxTemplateInspection,
  TemplatePlan,
  TemplatePlanErrorCode,
  TemplatePlanItem,
  TemplatePlanReport,
} from "../types";
import {
  buildMinimalDocx,
  buildSource,
  buildTemplate,
  fieldRuns,
  NS,
  p,
  partPaths,
  r,
  readPart,
  SOURCE_IDS,
  SOURCE_TEXT,
  TEMPLATE_IDS,
  XML,
} from "./fixtures";

const unwrap = <T>(result: Result<T, unknown>): T => {
  if (Result.isError(result)) {
    throw result.error;
  }
  return result.value;
};

const flatten = (blocks: readonly DocxBlock[]): DocxBlock[] =>
  blocks.flatMap((block) => [block, ...(block.nested ?? [])]);

type Fixture = {
  source: Uint8Array;
  template: Uint8Array;
  sourceManifest: DocxBlockManifest;
  templateInspection: DocxTemplateInspection;
  hash: (id: string) => string;
};

let fixture: Fixture;

beforeAll(async () => {
  const source = await buildSource();
  const template = await buildTemplate();
  const sourceManifest = unwrap(await analyzeDocxBlocks(source));
  const templateInspection = unwrap(await inspectDocxTemplate(template));
  const hashes = new Map(
    flatten(sourceManifest.blocks).map(({ id, hash }) => [id, hash]),
  );
  fixture = {
    source,
    template,
    sourceManifest,
    templateInspection,
    hash: (id) => hashes.get(id) ?? "",
  };
});

const sourceItem = (
  blockId: string,
  extra: Partial<Extract<TemplatePlanItem, { kind: "source" }>> = {},
): TemplatePlanItem => ({
  kind: "source",
  blockId,
  hash: fixture.hash(blockId),
  ...extra,
});

const spanOf = (
  text: string,
  needle: string,
): { start: number; end: number } => {
  const start = text.lastIndexOf(needle);
  if (start === -1) {
    throw new Error(`"${needle}" is not in "${text}"`);
  }
  return { start, end: start + needle.length };
};

/**
 * The plan the suite's main case applies: source numbering kept through a
 * numbering patch, the split headings joined, the manual label stripped, both
 * internal references linked, the template's fields frozen.
 */
const fullPlan = (): TemplatePlan => ({
  version: 1,
  sourceSha256: fixture.sourceManifest.sha256,
  templateSha256: fixture.templateInspection.sha256,
  body: [
    {
      kind: "slot",
      templateBlockId: TEMPLATE_IDS.title,
      sourceBlockIds: [SOURCE_IDS.title],
    },
    {
      kind: "template",
      templateBlockIds: [
        TEMPLATE_IDS.tocHeading,
        TEMPLATE_IDS.tocFirst,
        TEMPLATE_IDS.tocLast,
        TEMPLATE_IDS.fillIn,
      ],
    },
    sourceItem(SOURCE_IDS.articleOne, {
      styleId: "HouseArticle",
      stripLabel: "Article I",
      mergeNext: "none",
    }),
    sourceItem(SOURCE_IDS.articleOneTitle),
    sourceItem(SOURCE_IDS.clause, { styleId: "HouseClause" }),
    sourceItem(SOURCE_IDS.letterA, { styleId: "HouseLetter" }),
    sourceItem(SOURCE_IDS.letterB, { styleId: "HouseLetter" }),
    sourceItem(SOURCE_IDS.articleTwo, {
      styleId: "HouseArticle",
      stripLabel: "Article II",
      mergeNext: "none",
    }),
    sourceItem(SOURCE_IDS.articleTwoTitle),
    sourceItem(SOURCE_IDS.manualClause, {
      styleId: "HouseClause",
      stripLabel: "1.",
    }),
    sourceItem(SOURCE_IDS.link, { styleId: "HouseBody" }),
    sourceItem(SOURCE_IDS.table, {
      styleId: "HouseTable",
      cellStyleId: "HouseBody",
    }),
    sourceItem(SOURCE_IDS.mixed, { styleId: "HouseBody" }),
    { kind: "template", templateBlockIds: [TEMPLATE_IDS.signatures] },
  ],
  omit: [
    { blockId: SOURCE_IDS.blank, reason: "empty" },
    { blockId: SOURCE_IDS.pageBreak, reason: "pageArtifact" },
  ],
  numberingPatches: [
    { abstractNumId: 0, ilvl: 0, numFmt: "upperRoman", lvlText: "Article %1" },
    { abstractNumId: 0, ilvl: 1, lvlText: "%2." },
  ],
  references: [
    {
      blockId: SOURCE_IDS.clause,
      ...spanOf(SOURCE_TEXT.clause, "II"),
      targetBlockId: SOURCE_IDS.articleTwo,
      instruction: "\\n \\t \\h",
      expectedText: "II",
    },
    {
      blockId: SOURCE_IDS.manualClause,
      ...spanOf(SOURCE_TEXT.manualClause, "I"),
      targetBlockId: SOURCE_IDS.articleOne,
      instruction: "\\n \\t \\h",
      expectedText: "I",
    },
  ],
  highlights: [
    {
      blockId: SOURCE_IDS.link,
      ...spanOf(SOURCE_TEXT.link, "details"),
      color: "cyan",
      note: "unresolved",
    },
  ],
  fields: { freezeKinds: ["FILLIN", "DATE"], updateOnOpen: true },
});

type Applied = { docx: Uint8Array; report: TemplatePlanReport };

const apply = async (plan: TemplatePlan): Promise<Applied> =>
  unwrap(
    await applyTemplatePlan({
      source: fixture.source,
      template: fixture.template,
      plan,
    }),
  );

const refusal = async (
  plan: TemplatePlan,
  overrides: { source?: Uint8Array; template?: Uint8Array } = {},
): Promise<TemplatePlanErrorCode | "applied"> => {
  const result = await applyTemplatePlan({
    source: overrides.source ?? fixture.source,
    template: overrides.template ?? fixture.template,
    plan,
  });
  return Result.isError(result) ? result.error.code : "applied";
};

describe("applyTemplatePlan", () => {
  let applied: Applied;
  let output: DocxBlockManifest;
  let documentXml: string;

  beforeAll(async () => {
    applied = await apply(fullPlan());
    output = unwrap(await analyzeDocxBlocks(applied.docx));
    documentXml = (await readPart(applied.docx, "word/document.xml")) ?? "";
  });

  test("writes the planned story in order, labels rendered by the patched list", () => {
    expect(
      output.blocks.map(({ label, styleId, text }) => [label, styleId, text]),
    ).toEqual([
      [null, "HouseTitle", SOURCE_TEXT.title],
      [null, "HouseTitle", "Contents"],
      [null, "TOC1", "Sample article\t2"],
      [null, "TOC1", "Second sample\t3"],
      [null, "HouseBody", "To: Recipient name"],
      ["Article I", "HouseArticle", SOURCE_TEXT.articleOneTitle],
      ["1.", "HouseClause", SOURCE_TEXT.clause],
      ["a)", "HouseLetter", SOURCE_TEXT.letterA],
      ["b)", "HouseLetter", SOURCE_TEXT.letterB],
      ["Article II", "HouseArticle", SOURCE_TEXT.articleTwoTitle],
      [
        "1.",
        "HouseClause",
        "The fee follows from Article I of this agreement.",
      ],
      [null, "HouseBody", SOURCE_TEXT.link],
      [null, "HouseTable", "Item\tPrice\nService\t100"],
      [null, "HouseBody", SOURCE_TEXT.mixed],
      [
        null,
        "HouseTable",
        "........................\t........................",
      ],
    ]);
  });

  test("reports every source block identical", () => {
    const { report } = applied;

    expect(report.ok).toBe(true);
    expect(report.issues).toEqual([]);
    expect(report.counts).toEqual({
      sourceBlocks: 19,
      placed: 14,
      merged: 2,
      omitted: 3,
      templateBlocksKept: 5,
      textIdentical: 15,
      textMismatch: 0,
      emphasisMismatch: 0,
      labelChanged: 0,
      labelLost: 0,
    });
    expect(
      report.blocks.map(({ blockId, disposition, text }) => [
        blockId,
        disposition,
        text,
      ]),
    ).toEqual([
      ["p0", "slot", "identical"],
      ["p1", "omitted", "omitted"],
      ["p2", "placed", "identical"],
      ["p3", "merged", "identical"],
      ["p4", "placed", "identical"],
      ["p5", "placed", "identical"],
      ["p6", "placed", "identical"],
      ["p7", "placed", "identical"],
      ["p8", "merged", "identical"],
      ["p9", "placed", "identical"],
      ["p10", "placed", "identical"],
      ["t11", "placed", "identical"],
      ["p12", "nested", "identical"],
      ["p13", "nested", "identical"],
      ["p14", "nested", "identical"],
      ["p15", "nested", "identical"],
      ["p16", "placed", "identical"],
      ["p17", "omitted", "omitted"],
      ["p18", "omitted", "omitted"],
    ]);
    expect(report.blocks.find(({ blockId }) => blockId === "p9")).toMatchObject(
      { strippedPrefix: "1.\t", sourceLabel: "1.", outputLabel: "1." },
    );
    expect(report.blocks.find(({ blockId }) => blockId === "p2")).toMatchObject(
      {
        strippedPrefix: "Article I",
        sourceLabel: "Article I",
        outputLabel: "Article I",
      },
    );
    expect(
      report.blocks.find(({ blockId }) => blockId === "p17"),
    ).toMatchObject({ omitReason: "empty" });
  });

  test("carries the footnote over under a new id", async () => {
    expect(applied.report.footnotes).toEqual({
      source: 1,
      expected: 1,
      output: 1,
      identical: 1,
      mismatch: 0,
    });
    expect(output.footnotes).toEqual([{ id: "1", text: SOURCE_TEXT.footnote }]);
    const footnotes =
      (await readPart(applied.docx, "word/footnotes.xml")) ?? "";
    expect(footnotes).toContain('<w:pStyle w:val="FootnoteText"/>');
    expect(footnotes).toContain("<w:i/><w:iCs/>");
    expect(documentXml).toContain(
      '<w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="1"/>',
    );
  });

  test("keeps stated emphasis and highlight, and drops fonts and colours", () => {
    const { report } = applied;

    // Source: bold title, italic word, footnote italic, bold table header
    // cells, Strong run, underline and subscript; one yellow highlight.
    expect(report.emphasisSpans).toEqual({ source: 8, expected: 8, output: 8 });
    // The plan adds a cyan highlight of its own.
    expect(report.highlightSpans).toEqual({
      source: 1,
      expected: 2,
      output: 2,
    });
    expect(documentXml).toContain(
      '<w:r><w:rPr><w:i/><w:iCs/></w:rPr><w:t xml:space="preserve">services</w:t></w:r>',
    );
    expect(documentXml).toContain(
      '<w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t xml:space="preserve">[to be completed]</w:t>',
    );
    expect(documentXml).toContain(
      '<w:rPr><w:highlight w:val="cyan"/></w:rPr><w:t xml:space="preserve">details</w:t>',
    );
    // The character style's bold is stated on the run; the style id is not.
    expect(documentXml).toContain(
      '<w:rPr><w:b/><w:bCs/></w:rPr><w:t xml:space="preserve">this agreement</w:t>',
    );
    expect(documentXml).toContain('<w:vertAlign w:val="subscript"/>');
    expect(documentXml).not.toContain("Times New Roman");
    expect(documentXml).not.toContain("FF0000");
    expect(documentXml).not.toContain('w:val="Strong"');
    expect(documentXml).not.toContain("<w:jc");
    expect(documentXml).toContain(" described in Article ");
    expect(documentXml).toContain("<w:tab/>");
    expect(documentXml).toContain("<w:softHyphen/>");
  });

  test("writes references as REF fields on hidden bookmarks", () => {
    const { references } = applied.report;

    expect(references).toMatchObject({
      planned: 2,
      ok: 2,
      textChanging: 0,
      predictionMismatch: 0,
    });
    expect(
      references.results.map(
        ({ blockId, targetBlockId, cachedText, predictedText, status }) => [
          blockId,
          targetBlockId,
          cachedText,
          predictedText,
          status,
        ],
      ),
    ).toEqual([
      ["p4", "p7", "II", "II", "ok"],
      ["p9", "p2", "I", "I", "ok"],
    ]);
    for (const { bookmark } of references.results) {
      expect(bookmark).toMatch(/^_Ref\d{9}$/u);
      expect(documentXml).toContain(
        `<w:instrText xml:space="preserve"> REF ${bookmark} \\n \\t \\h </w:instrText>`,
      );
      expect(documentXml.split(`w:name="${bookmark}"`)).toHaveLength(2);
    }
    // The bookmark wraps the heading's content, after its properties.
    expect(documentXml).toMatch(
      /<w:pPr><w:pStyle w:val="HouseArticle"\/><\/w:pPr><w:bookmarkStart w:id="\d+" w:name="_Ref\d{9}"\/><w:r><w:t xml:space="preserve">Fees<\/w:t><\/w:r><w:bookmarkEnd w:id="\d+"\/><\/w:p>/u,
    );
  });

  test("re-homes the hyperlink and applies the template's link style", async () => {
    const relationships =
      (await readPart(applied.docx, "word/_rels/document.xml.rels")) ?? "";
    const added =
      /Id="(?<id>rId\d+)"[^>]*Target="https:\/\/example\.com\/register\?a=1&amp;b=2" TargetMode="External"/u.exec(
        relationships,
      )?.groups?.["id"];

    expect(added).toBeDefined();
    expect(documentXml).toContain(
      `<w:hyperlink r:id="${added}" w:history="1"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t xml:space="preserve">public register</w:t></w:r></w:hyperlink>`,
    );
  });

  test("rebuilds the table on the template's table style", () => {
    expect(documentXml).toContain(
      '<w:tbl><w:tblPr><w:tblStyle w:val="HouseTable"/><w:tblW w:w="5000" w:type="pct"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:trPr><w:tblHeader/></w:trPr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:pStyle w:val="HouseBody"/></w:pPr><w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t xml:space="preserve">Item</w:t></w:r></w:p></w:tc>',
    );
    expect(documentXml).not.toContain("D9D9D9");
  });

  test("patches the numbering in place and leaves the rest of the part alone", async () => {
    const before =
      (await readPart(fixture.template, "word/numbering.xml")) ?? "";
    const after = (await readPart(applied.docx, "word/numbering.xml")) ?? "";

    expect(after).toBe(
      before
        .replace(
          '<w:numFmt w:val="decimal"/>',
          '<w:numFmt w:val="upperRoman"/>',
        )
        .replace(
          '<w:lvlText w:val="Art. %1"/>',
          '<w:lvlText w:val="Article %1"/>',
        )
        .replace('<w:lvlText w:val="%1.%2."/>', '<w:lvlText w:val="%2."/>'),
    );
  });

  test("freezes the planned field kinds and sets update on open", async () => {
    const { report } = applied;
    const footer = (await readPart(applied.docx, "word/footer1.xml")) ?? "";
    const settings = (await readPart(applied.docx, "word/settings.xml")) ?? "";

    expect(footer).toContain("1.1.2020");
    expect(footer).not.toContain("DATE");
    expect(footer).toContain("FILENAME");
    expect(documentXml).not.toContain("FILLIN");
    expect(documentXml).toContain("Recipient name");
    expect(settings).toMatch(
      /<w:listSeparator w:val=";"\/><\/w:settings>$|updateFields/u,
    );
    expect(settings.indexOf("updateFields")).toBeGreaterThan(
      settings.indexOf("characterSpacingControl"),
    );
    expect(settings.indexOf("updateFields")).toBeLessThan(
      settings.indexOf("<w:compat>"),
    );
    expect(report.updateFields).toBe("set");
    expect(report.leftovers.frozenFields).toEqual([
      { part: "word/document.xml", kind: "FILLIN", count: 1 },
      { part: "word/footer1.xml", kind: "DATE", count: 1 },
    ]);
    expect(report.leftovers.fields).toEqual([
      { part: "word/document.xml", kind: "TOC", count: 1 },
      { part: "word/document.xml", kind: "PAGEREF", count: 2 },
      { part: "word/header1.xml", kind: "PAGE", count: 1 },
      { part: "word/footer1.xml", kind: "FILENAME", count: 1 },
    ]);
    expect(report.leftovers.templateBlockIds).toEqual([
      TEMPLATE_IDS.tocHeading,
      TEMPLATE_IDS.tocFirst,
      TEMPLATE_IDS.tocLast,
      TEMPLATE_IDS.fillIn,
      TEMPLATE_IDS.signatures,
    ]);
  });

  test("reports what the transfer drops and warns about stale template fields", () => {
    const { report } = applied;

    expect(report.dropped).toEqual({
      comments: 1,
      sourceFields: [],
      internalHyperlinks: 0,
      sectionBreaks: 0,
    });
    expect(report.conformance).toMatchObject({
      status: "conformant",
      templateStatus: "conformant",
    });
    expect(report.warnings.join("\n")).toContain("comment(s) of the source");
    expect(report.warnings.join("\n")).toContain("table of contents");
    expect(report.warnings.join("\n")).toContain("_Toc100000001");
    expect(documentXml).not.toContain("commentReference");
  });

  test("keeps untouched template parts byte-identical", async () => {
    const before = await ZipArchive.read(fixture.template);
    const after = await ZipArchive.read(applied.docx);

    for (const path of ["word/styles.xml", "word/header1.xml", "_rels/.rels"]) {
      // oxlint-disable-next-line no-await-in-loop -- three small reads
      expect(await after.entry(path)?.bytes()).toEqual(
        // oxlint-disable-next-line no-await-in-loop -- three small reads
        await before.entry(path)?.bytes(),
      );
    }
    expect(await partPaths(applied.docx)).toEqual(
      await partPaths(fixture.template),
    );
    // The template's final section properties close the new body.
    expect(documentXml).toMatch(
      /<w:sectPr><w:headerReference w:type="default" r:id="rId10"\/>.*<\/w:sectPr><\/w:body><\/w:document>$/u,
    );
  });

  test("is deterministic", async () => {
    const again = await apply(fullPlan());

    expect(again.report.outputSha256).toBe(applied.report.outputSha256);
  });

  test("counts filler strings the caller names", async () => {
    const result = unwrap(
      await applyTemplatePlan({
        source: fixture.source,
        template: fixture.template,
        plan: fullPlan(),
        options: {
          fillerStrings: ["........................", "[party name]"],
        },
      }),
    );

    expect(result.report.leftovers.filler).toEqual([
      { text: "........................", part: "word/document.xml", count: 2 },
    ]);
  });
});

describe("applyTemplatePlan numbering", () => {
  test("starts a new list instance for a group and cancels a style's list", async () => {
    const plan = fullPlan();
    plan.body = plan.body.map((item) => {
      if (item.kind !== "source") {
        return item;
      }
      if (
        item.blockId === SOURCE_IDS.letterA ||
        item.blockId === SOURCE_IDS.letterB
      ) {
        return {
          ...item,
          numbering: {
            numId: "new",
            ilvl: 3,
            restartFrom: 3,
            group: "letters",
          },
        };
      }
      return item.blockId === SOURCE_IDS.manualClause
        ? { ...item, numbering: null }
        : item;
    });
    plan.references = [];
    const { docx, report } = await apply(plan);
    const numbering = (await readPart(docx, "word/numbering.xml")) ?? "";
    const document = (await readPart(docx, "word/document.xml")) ?? "";
    const blocks = unwrap(await analyzeDocxBlocks(docx)).blocks;

    expect(numbering).toContain(
      '<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num><w:num w:numId="3"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="3"><w:startOverride w:val="3"/></w:lvlOverride></w:num></w:numbering>',
    );
    expect(
      document.match(/<w:numPr><w:ilvl w:val="3"\/><w:numId w:val="3"\/>/gu),
    ).toHaveLength(2);
    expect(document).toContain('<w:numPr><w:numId w:val="0"/></w:numPr>');
    expect(blocks.map(({ label }) => label).filter(Boolean)).toEqual([
      "Article I",
      "1.",
      "c)",
      "d)",
      "Article II",
    ]);
    // The source showed "a)", "b)" and a manual "1."; the report says so.
    expect(report.counts).toMatchObject({ labelChanged: 2, labelLost: 1 });
    expect(report.ok).toBe(true);
  });

  test("refuses a list the template does not define", async () => {
    const plan = fullPlan();
    plan.body = plan.body.map((item) =>
      item.kind === "source" && item.blockId === SOURCE_IDS.link
        ? { ...item, numbering: { numId: 42, ilvl: 0 } }
        : item,
    );

    expect(await refusal(plan)).toBe("numbering-unknown");
    expect(
      await refusal({
        ...fullPlan(),
        numberingPatches: [{ abstractNumId: 9, ilvl: 0, lvlText: "%1" }],
      }),
    ).toBe("numbering-unknown");
  });
});

describe("applyTemplatePlan refusals", () => {
  const withBody = (
    edit: (body: TemplatePlanItem[]) => TemplatePlanItem[],
  ): TemplatePlan => {
    const plan = fullPlan();
    return { ...plan, body: edit(plan.body) };
  };

  test("refuses a plan made for other bytes", async () => {
    expect(await refusal({ ...fullPlan(), sourceSha256: "0".repeat(64) })).toBe(
      "sha-mismatch",
    );
    expect(
      await refusal({ ...fullPlan(), templateSha256: "0".repeat(64) }),
    ).toBe("sha-mismatch");
  });

  test("refuses unknown ids and stale hashes", async () => {
    expect(
      await refusal(
        withBody((body) => [
          ...body,
          { kind: "source", blockId: "p99", hash: "x" },
        ]),
      ),
    ).toBe("unknown-block");
    expect(
      await refusal(
        withBody((body) => [
          ...body,
          { kind: "template", templateBlockIds: ["p99"] },
        ]),
      ),
    ).toBe("unknown-block");
    // A paragraph inside a table travels with the table.
    expect(
      await refusal(withBody((body) => [...body, sourceItem("p12")])),
    ).toBe("unknown-block");
    expect(
      await refusal(
        withBody((body) =>
          body.map((item) =>
            item.kind === "source" && item.blockId === SOURCE_IDS.link
              ? { ...item, hash: "000000000000" }
              : item,
          ),
        ),
      ),
    ).toBe("hash-mismatch");
  });

  test("refuses a dropped block and a block placed twice", async () => {
    expect(
      await refusal(
        withBody((body) =>
          body.filter(
            (item) =>
              item.kind !== "source" || item.blockId !== SOURCE_IDS.link,
          ),
        ),
      ),
    ).toBe("block-unplaced");
    expect(
      await refusal(withBody((body) => [...body, sourceItem(SOURCE_IDS.link)])),
    ).toBe("block-placed-twice");
    expect(
      await refusal({
        ...fullPlan(),
        omit: [
          ...fullPlan().omit,
          { blockId: SOURCE_IDS.link, reason: "sourceToc" },
        ],
      }),
    ).toBe("block-placed-twice");
    // "empty" is checked, not taken on trust.
    const dropped = withBody((body) =>
      body.filter(
        (item) => item.kind !== "source" || item.blockId !== SOURCE_IDS.link,
      ),
    );
    expect(
      await refusal({
        ...dropped,
        omit: [...dropped.omit, { blockId: SOURCE_IDS.link, reason: "empty" }],
      }),
    ).toBe("plan-invalid");
  });

  test("refuses a label the block does not start with", async () => {
    expect(
      await refusal(
        withBody((body) =>
          body.map((item) =>
            item.kind === "source" && item.blockId === SOURCE_IDS.manualClause
              ? { ...item, stripLabel: "2." }
              : item,
          ),
        ),
      ),
    ).toBe("strip-label-mismatch");
  });

  test("refuses a merge that has nothing to join", async () => {
    expect(
      await refusal(
        withBody((body) =>
          body.map((item) =>
            item.kind === "source" && item.blockId === SOURCE_IDS.mixed
              ? { ...item, mergeNext: "space" }
              : item,
          ),
        ),
      ),
    ).toBe("merge-invalid");
  });

  test("refuses reference spans outside the block and unplaced targets", async () => {
    const plan = fullPlan();
    const [first] = plan.references ?? [];
    if (first === undefined) {
      throw new Error("the plan has no reference");
    }

    expect(
      await refusal({ ...plan, references: [{ ...first, end: 999 }] }),
    ).toBe("reference-span-invalid");
    expect(
      await refusal({
        ...plan,
        references: [{ ...first, targetBlockId: SOURCE_IDS.blank }],
      }),
    ).toBe("reference-target-unplaced");
    // The stripped label is gone; nothing in it can become a field.
    expect(
      await refusal({
        ...plan,
        references: [
          {
            ...first,
            blockId: SOURCE_IDS.manualClause,
            start: 0,
            end: 1,
          },
        ],
      }),
    ).toBe("reference-span-invalid");
    // A span over a tab is not text a field can cache.
    expect(
      await refusal({
        ...plan,
        references: [
          {
            ...first,
            blockId: SOURCE_IDS.letterA,
            ...spanOf(SOURCE_TEXT.letterA, "item\twith"),
          },
        ],
      }),
    ).toBe("reference-span-invalid");
    expect(
      await refusal({
        ...plan,
        references: [{ ...first, instruction: '\\h "free text"' }],
      }),
    ).toBe("plan-invalid");
    expect(
      await refusal({
        ...plan,
        highlights: [
          { blockId: SOURCE_IDS.link, start: 0, end: 3, color: "#ff0000" },
        ],
      }),
    ).toBe("highlight-span-invalid");
  });

  test("refuses a style the template does not have", async () => {
    expect(
      await refusal(
        withBody((body) =>
          body.map((item) =>
            item.kind === "source" && item.blockId === SOURCE_IDS.link
              ? { ...item, styleId: "Heading1" }
              : item,
          ),
        ),
      ),
    ).toBe("style-unknown");
  });

  test("refuses a source with tracked changes", async () => {
    const tracked = await buildMinimalDocx(
      p(
        `${r("kept ")}<w:ins w:id="1" w:author="a" w:date="2020-01-01T00:00:00Z">${r("added")}</w:ins>`,
      ),
    );
    const manifest = unwrap(await analyzeDocxBlocks(tracked));
    const block = manifest.blocks.at(0);

    expect(
      await refusal(
        {
          version: 1,
          sourceSha256: manifest.sha256,
          templateSha256: fixture.templateInspection.sha256,
          body: [{ kind: "source", blockId: "p0", hash: block?.hash ?? "" }],
          omit: [],
        },
        { source: tracked },
      ),
    ).toBe("tracked-changes");
  });

  test("refuses a template field kept in part", async () => {
    const result = await applyTemplatePlan({
      source: fixture.source,
      template: fixture.template,
      plan: withBody((body) =>
        body.map((item) =>
          item.kind === "template" &&
          item.templateBlockIds.includes(TEMPLATE_IDS.tocFirst)
            ? {
                ...item,
                templateBlockIds: item.templateBlockIds.filter(
                  (id) => id !== TEMPLATE_IDS.tocLast,
                ),
              }
            : item,
        ),
      ),
    });

    expect(Result.isError(result)).toBe(true);
    if (Result.isError(result)) {
      expect(result.error.code).toBe("verification-failed");
      expect(result.error.report?.ok).toBe(false);
      expect(result.error.report?.issues.map(({ code }) => code)).toContain(
        "field-unbalanced",
      );
    }
  });

  test("refuses a prompting field left under update on open", async () => {
    const plan = fullPlan();
    const result = await applyTemplatePlan({
      source: fixture.source,
      template: fixture.template,
      plan: { ...plan, fields: { freezeKinds: ["DATE"], updateOnOpen: true } },
    });

    expect(Result.isError(result)).toBe(true);
    if (Result.isError(result)) {
      expect(result.error.report?.issues.map(({ code }) => code)).toEqual([
        "prompting-field-left",
      ]);
    }
  });
});

describe("applyTemplatePlan onto a template without notes", () => {
  test("creates the footnotes part with its relationship and content type", async () => {
    const template = await buildTemplate({ withoutFootnotes: true });
    const inspection = unwrap(await inspectDocxTemplate(template));
    const result = unwrap(
      await applyTemplatePlan({
        source: fixture.source,
        template,
        plan: { ...fullPlan(), templateSha256: inspection.sha256 },
      }),
    );
    const footnotes = (await readPart(result.docx, "word/footnotes.xml")) ?? "";

    expect(result.report.footnotes).toMatchObject({
      identical: 1,
      mismatch: 0,
    });
    expect(footnotes).toContain('w:type="separator" w:id="-1"');
    expect(footnotes).toContain('<w:footnote w:id="1">');
    expect(await readPart(result.docx, "[Content_Types].xml")).toContain(
      '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>',
    );
    expect(await readPart(result.docx, "word/_rels/document.xml.rels")).toMatch(
      /relationships\/footnotes" Target="footnotes\.xml"/u,
    );
  });
});

const REL_NS =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PIXEL_PNG = Uint8Array.from(
  Buffer.from(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==",
    "base64",
  ),
);
const PICTURE_RUN =
  '<w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" distT="0" distB="0" distL="0" distR="0"><wp:extent cx="100000" cy="100000"/><wp:docPr id="7" name="Picture 7"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="0" name="image1.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId5"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100000" cy="100000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';

type ObjectSourceOptions = { imageRelationshipType?: string };

/** A source with a picture, fields, an internal link and an endnote. */
const buildObjectSource = async ({
  imageRelationshipType = "image",
}: ObjectSourceOptions = {}): Promise<Uint8Array> =>
  await buildMinimalDocx(
    [
      p(`${r("Picture: ")}${PICTURE_RUN}`),
      p(
        `${r("See ")}${fieldRuns('HYPERLINK "https://example.org/site"', r("the site"))}${r(" and ")}<w:fldSimple w:instr=" REF _Ref1 \\h ">${r("clause 3")}</w:fldSimple>${r(" or ")}<w:hyperlink w:anchor="_Ref1">${r("above")}</w:hyperlink>`,
      ),
      p(`${r("End")}<w:r><w:endnoteReference w:id="1"/></w:r>`),
    ].join(""),
    {
      "word/_rels/document.xml.rels": `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId5" Type="${REL_NS}/${imageRelationshipType}" Target="media/image1.png"/><Relationship Id="rId6" Type="${REL_NS}/endnotes" Target="endnotes.xml"/></Relationships>`,
      "word/media/image1.png": PIXEL_PNG,
      "word/endnotes.xml": `${XML}<w:endnotes ${NS}><w:endnote w:id="1"><w:p><w:r><w:endnoteRef/></w:r>${r(" An endnote.")}</w:p></w:endnote></w:endnotes>`,
    },
  );

describe("applyTemplatePlan with objects, fields and endnotes", () => {
  const planFor = async (source: Uint8Array): Promise<TemplatePlan> => {
    const manifest = unwrap(await analyzeDocxBlocks(source));
    return {
      version: 1,
      sourceSha256: manifest.sha256,
      templateSha256: fixture.templateInspection.sha256,
      body: manifest.blocks.map(({ id, hash }) => ({
        kind: "source",
        blockId: id,
        hash,
        styleId: "HouseBody",
      })),
      omit: [],
    };
  };

  test("moves a picture with its media, relationship and content type", async () => {
    const source = await buildObjectSource();
    const { docx, report } = unwrap(
      await applyTemplatePlan({
        source,
        template: fixture.template,
        plan: await planFor(source),
      }),
    );
    const document = (await readPart(docx, "word/document.xml")) ?? "";
    const relationships =
      (await readPart(docx, "word/_rels/document.xml.rels")) ?? "";
    const imageId =
      /Id="(?<id>rId\d+)"[^>]*relationships\/image" Target="\/word\/media\/placed1\.png"/u.exec(
        relationships,
      )?.groups?.["id"];

    expect(report.ok).toBe(true);
    expect(imageId).toBeDefined();
    expect(document).toContain(`<a:blip r:embed="${imageId}"/>`);
    expect(document).toContain('<wp:docPr id="1" name="Picture 7"/>');
    expect(
      await (
        await ZipArchive.read(docx)
      )
        .entry("word/media/placed1.png")
        ?.bytes(),
    ).toEqual(PIXEL_PNG);
    expect(await readPart(docx, "[Content_Types].xml")).toContain(
      '<Override PartName="/word/media/placed1.png" ContentType="image/png"/>',
    );
    expect(
      unwrap(await analyzeDocxBlocks(docx)).blocks.at(0)?.inlineObjects,
    ).toBe(1);
  });

  test("keeps field results as text and a HYPERLINK field as a link", async () => {
    const source = await buildObjectSource();
    const { docx, report } = unwrap(
      await applyTemplatePlan({
        source,
        template: fixture.template,
        plan: await planFor(source),
      }),
    );
    const document = (await readPart(docx, "word/document.xml")) ?? "";
    const relationships =
      (await readPart(docx, "word/_rels/document.xml.rels")) ?? "";

    expect(unwrap(await analyzeDocxBlocks(docx)).blocks.at(1)?.text).toBe(
      "See the site and clause 3 or above",
    );
    expect(document).not.toContain("fldChar");
    expect(document).not.toContain("fldSimple");
    expect(relationships).toContain(
      'Target="https://example.org/site" TargetMode="External"',
    );
    expect(document).toMatch(
      /<w:hyperlink r:id="rId\d+" w:history="1"><w:r><w:rPr><w:rStyle w:val="Hyperlink"\/><\/w:rPr><w:t xml:space="preserve">the site<\/w:t><\/w:r><\/w:hyperlink>/u,
    );
    expect(report.dropped).toMatchObject({
      sourceFields: [
        { kind: "HYPERLINK", count: 1 },
        { kind: "REF", count: 1 },
      ],
      internalHyperlinks: 1,
    });
  });

  test("creates the endnotes part for a template that has none", async () => {
    const source = await buildObjectSource();
    const { docx, report } = unwrap(
      await applyTemplatePlan({
        source,
        template: fixture.template,
        plan: await planFor(source),
      }),
    );

    expect(report.endnotes).toEqual({
      source: 1,
      expected: 1,
      output: 1,
      identical: 1,
      mismatch: 0,
    });
    expect(await readPart(docx, "word/endnotes.xml")).toContain(
      '<w:endnote w:id="1"><w:p><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteRef/></w:r>',
    );
    expect(unwrap(await analyzeDocxBlocks(docx)).endnotes).toEqual([
      { id: "1", text: " An endnote." },
    ]);
  });

  test("refuses an embedded object that is not a picture", async () => {
    const source = await buildObjectSource({ imageRelationshipType: "chart" });

    expect(await refusal(await planFor(source), { source })).toBe(
      "unsupported-content",
    );
  });
});
