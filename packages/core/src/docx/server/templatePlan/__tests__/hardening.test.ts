import { Result } from "better-result";
import { beforeAll, describe, expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { applyTemplatePlan } from "../apply";
import { analyzeDocxBlocks } from "../blocks";
import {
  keepsShiftReturnSpacing,
  sectionPropertiesXml,
  setKeepShiftReturnSpacing,
} from "../builders";
import { isManualLabel } from "../plan";
import type {
  TemplatePlan,
  TemplatePlanItem,
  TemplatePlanSourceItem,
} from "../types";
import { readRawPart } from "../xml";
import {
  buildMinimalDocx,
  buildSource,
  buildTemplate,
  fieldRuns,
  NS,
  p,
  r,
  readPart,
  replacePart,
  style,
  TEMPLATE_IDS,
  XML,
} from "./fixtures";
import {
  applyFor,
  buildSourceWith,
  codeOf,
  issuesOf,
  pictureRun,
  PIXEL_PNG,
  planFor,
  type PlannedBlock,
  revision,
  unwrap,
  verifyFor,
} from "./planning";

let template: Uint8Array;

beforeAll(async () => {
  template = await buildTemplate();
});

const SEPARATORS =
  '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>';
const NOTE_REFERENCE = '<w:r><w:footnoteReference w:id="2"/></w:r>';
const note = (inner: string): string =>
  `${SEPARATORS}<w:footnote w:id="2">${inner}</w:footnote>`;

const styled =
  (styleId: string, extra: Partial<TemplatePlanSourceItem> = {}) =>
  (blocks: PlannedBlock[]): TemplatePlanItem[] =>
    blocks.map(({ id, hash }) => ({
      kind: "source",
      blockId: id,
      hash,
      styleId,
      ...extra,
    }));

const documentOf = async (docx: Uint8Array): Promise<string> =>
  (await readPart(docx, "word/document.xml")) ?? "";

describe("highlights and note references", () => {
  test("a highlight over a note reference applies and leaves the mark plain", async () => {
    const source = await buildSourceWith({
      body: p(`${r("See clause text")}${NOTE_REFERENCE}${r(" below.")}`),
      footnotes: note(p(r("The note."))),
    });
    const applied = await applyFor({
      source,
      template,
      extra: {
        highlights: [{ blockId: "p0", start: 0, end: 23, color: "yellow" }],
      },
    });
    const { report } = unwrap(applied);

    expect(report.ok).toBe(true);
    expect(report.highlightSpans).toMatchObject({ expected: 2, output: 2 });
    expect(report.footnotes).toMatchObject({ identical: 1, mismatch: 0 });
  });

  test("overlapping highlights are refused", async () => {
    const source = await buildMinimalDocx(p(r("Some plain text here.")));

    expect(
      codeOf(
        await applyFor({
          source,
          template,
          extra: {
            highlights: [
              { blockId: "p0", start: 0, end: 9, color: "yellow" },
              { blockId: "p0", start: 5, end: 12, color: "green" },
            ],
          },
        }),
      ),
    ).toBe("highlight-span-invalid");
  });
});

describe("pictures", () => {
  const OTHER_PNG = Uint8Array.from([...PIXEL_PNG, 0]);
  const twoPictures = async (): Promise<Uint8Array> =>
    await buildSourceWith({
      body: `${p(`${r("One ")}${pictureRun("rId5", 1)}`)}${p(`${r("Two ")}${pictureRun("rId6", 2)}`)}`,
      relationships: [
        { id: "rId5", type: "image", target: "media/image1.png" },
        { id: "rId6", type: "image", target: "media/image2.png" },
      ],
      parts: {
        "word/media/image1.png": PIXEL_PNG,
        "word/media/image2.png": OTHER_PNG,
      },
    });
  const mediaOf = async (
    docx: Uint8Array,
    path: string,
  ): Promise<Uint8Array | undefined> =>
    await (await ZipArchive.read(docx)).entry(path)?.bytes();

  test("two pictures get two parts when the template already holds the first name", async () => {
    const taken = await ZipArchive.read(template);
    taken.set("word/media/placed1.png", PIXEL_PNG);
    const crowded = await taken.generate();
    const source = await twoPictures();
    const { docx, report } = unwrap(
      await applyFor({ source, template: crowded }),
    );

    expect(report.ok).toBe(true);
    expect(await mediaOf(docx, "word/media/placed2.png")).toEqual(PIXEL_PNG);
    expect(await mediaOf(docx, "word/media/placed3.png")).toEqual(OTHER_PNG);
  });

  test("verification notices a picture whose bytes are another picture's", async () => {
    const source = await twoPictures();
    const plan = await planFor({ source, template });
    const { docx } = unwrap(
      await applyTemplatePlan({ source, template, plan }),
    );
    const swapped = await ZipArchive.read(docx);
    swapped.set("word/media/placed1.png", OTHER_PNG);
    const report = await verifyFor({
      source,
      template,
      output: await swapped.generate(),
      plan,
    });

    expect(report.ok).toBe(false);
    expect(report.issues.map(({ code, blockId }) => [code, blockId])).toEqual([
      ["text-mismatch", "p0"],
    ]);
  });

  test("a percent-encoded target finds its picture", async () => {
    const source = await buildSourceWith({
      body: p(`${r("Exhibit ")}${pictureRun("rId5")}`),
      relationships: [
        { id: "rId5", type: "image", target: "media/Exhibit%20A.png" },
      ],
      parts: { "word/media/Exhibit A.png": PIXEL_PNG },
    });
    const { docx, report } = unwrap(await applyFor({ source, template }));

    expect(report.ok).toBe(true);
    expect(await mediaOf(docx, "word/media/placed1.png")).toEqual(PIXEL_PNG);
  });
});

describe("tracked changes outside paragraph text", () => {
  const cell = (inner: string, properties = ""): string =>
    `<w:tc>${properties === "" ? "" : `<w:tcPr>${properties}</w:tcPr>`}${inner}</w:tc>`;
  const cases: [string, string][] = [
    [
      "a deleted table row",
      `<w:tbl><w:tr><w:trPr><w:del ${revision(1)}/></w:trPr>${cell(p(r("gone")))}</w:tr><w:tr>${cell(p(r("kept")))}</w:tr></w:tbl>`,
    ],
    [
      "a deleted table cell",
      `<w:tbl><w:tr>${cell(p(r("gone")), `<w:cellDel ${revision(1)}/>`)}${cell(p(r("kept")))}</w:tr></w:tbl>`,
    ],
    [
      "a block-level insertion",
      `${p(r("kept"))}<w:ins ${revision(1)}>${r("stray")}</w:ins>`,
    ],
    [
      "an insertion inside a text box",
      p(
        `${r("before ")}<w:r><w:pict><v:shape xmlns:v="urn:schemas-microsoft-com:vml"><v:textbox><w:txbxContent><w:p><w:ins ${revision(1)}>${r("boxed")}</w:ins></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r>`,
      ),
    ],
  ];

  test.each(cases)("%s is counted and refused", async (_name, body) => {
    const source = await buildMinimalDocx(body);

    expect(
      unwrap(await analyzeDocxBlocks(source)).trackedChanges,
    ).toBeGreaterThan(0);
    expect(codeOf(await applyFor({ source, template }))).toBe(
      "tracked-changes",
    );
  });
});

describe("content the transfer has no reading for", () => {
  const MC =
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"';
  const cases: [string, string][] = [
    [
      "an unclosed field",
      `${p(`${r("Start ")}<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> REF x </w:instrText></w:r>`)}${p(r("An important clause that follows."))}`,
    ],
    [
      "ruby text",
      p(
        `<w:r><w:ruby><w:rt><w:r><w:t>kana</w:t></w:r></w:rt><w:rubyBase><w:r><w:t>kanji</w:t></w:r></w:rubyBase></w:ruby></w:r>`,
      ),
    ],
    [
      "alternate content beside runs",
      p(
        `${r("old or ")}<mc:AlternateContent ${MC}><mc:Choice Requires="w14">${r("new")}</mc:Choice><mc:Fallback>${r("old")}</mc:Fallback></mc:AlternateContent>`,
      ),
    ],
    [
      "an embedded chunk between blocks",
      `${p(r("before"))}<w:altChunk r:id="rId9"/>${p(r("after"))}`,
    ],
  ];

  test.each(cases)("%s is refused, not dropped", async (_name, body) => {
    const source = await buildMinimalDocx(body);

    expect(codeOf(await applyFor({ source, template }))).toBe(
      "unsupported-content",
    );
  });

  test("an unclosed field is named in the manifest", async () => {
    const source = await buildMinimalDocx(
      `${p(`<w:r><w:fldChar w:fldCharType="begin"/></w:r>`)}${p(r("Hidden by the field."))}`,
    );

    expect(
      unwrap(await analyzeDocxBlocks(source)).warnings.join(" "),
    ).toContain("do not pair up");
  });

  test("a note with a table is refused", async () => {
    const source = await buildSourceWith({
      body: p(`${r("Text")}${NOTE_REFERENCE}`),
      footnotes: note(
        `${p(r("Lead."))}<w:tbl><w:tr><w:tc>${p(r("cell"))}</w:tc></w:tr></w:tbl>`,
      ),
    });

    expect(codeOf(await applyFor({ source, template }))).toBe(
      "unsupported-content",
    );
  });

  test("a reference to a note the source does not hold is refused", async () => {
    const source = await buildSourceWith({
      body: p(`${r("Text")}<w:r><w:footnoteReference w:id="7"/></w:r>`),
      footnotes: SEPARATORS,
    });

    expect(codeOf(await applyFor({ source, template }))).toBe(
      "unsupported-content",
    );
  });

  test("a notes part that cannot be read fails the analysis", async () => {
    const source = await buildSourceWith({
      body: p(`${r("Text")}${NOTE_REFERENCE}`),
      footnotes: note(p(r("The note."))),
    });
    const broken = await replacePart(
      source,
      "word/footnotes.xml",
      `<w:footnotes ${NS}><w:footnote w:id="2"><w:p>`,
    );
    const analysis = await analyzeDocxBlocks(broken);

    expect(Result.isError(analysis)).toBe(true);
    if (Result.isError(analysis)) {
      expect(analysis.error.code).toBe("xml-unreadable");
    }
  });
});

describe("packages outside the supported profile", () => {
  const STRICT = "http://purl.oclc.org/ooxml/wordprocessingml/main";
  const TRANSITIONAL =
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

  test("a Strict source is refused up front", async () => {
    const plain = await buildMinimalDocx(p(r("Body text.")));
    const xml = (await readPart(plain, "word/document.xml")) ?? "";
    const strict = await replacePart(
      plain,
      "word/document.xml",
      xml.replace(TRANSITIONAL, STRICT),
    );

    expect(codeOf(await applyFor({ source: strict, template }))).toBe(
      "unsupported-content",
    );
  });

  test("a template that binds no prefix to its markup is refused, not thrown at", async () => {
    const xml = (await readPart(template, "word/document.xml")) ?? "";
    const unprefixed = await replacePart(
      template,
      "word/document.xml",
      `${XML}<document xmlns="${TRANSITIONAL}"><body><p><r><t>Only</t></r></p><sectPr/></body></document>`,
    );
    expect(xml).not.toBe("");
    const source = await buildMinimalDocx(p(r("Body text.")));

    expect(codeOf(await applyFor({ source, template: unprefixed }))).toBe(
      "xml-unreadable",
    );
  });

  test("a template whose content-type registry is not one returns an error", async () => {
    const bare = await buildTemplate({ withoutFootnotes: true });
    const broken = await replacePart(
      bare,
      "[Content_Types].xml",
      `${XML}<Wrong xmlns="urn:example"/>`,
    );
    const source = await buildSourceWith({
      body: p(`${r("Text")}${NOTE_REFERENCE}`),
      footnotes: note(p(r("The note."))),
    });
    const applied = await applyFor({ source, template: broken });

    expect(codeOf(applied)).toBe("xml-unreadable");
  });
});

describe("malformed plans", () => {
  let source: Uint8Array;
  let valid: TemplatePlan;

  beforeAll(async () => {
    source = await buildMinimalDocx(
      `${p(r("First paragraph."))}${p(r("Second paragraph."))}`,
    );
    valid = await planFor({ source, template });
  });

  // A plan arrives as parsed JSON, so the cases are JSON too.
  type Json = Record<string, unknown>;
  const isJson = (value: unknown): value is Json =>
    typeof value === "object" && value !== null && !Array.isArray(value);
  const mutated = (edit: (plan: Json) => unknown): unknown => {
    const copy: unknown = JSON.parse(JSON.stringify(valid));
    if (!isJson(copy)) {
      throw new Error("the fixture plan is not an object");
    }
    return edit(copy);
  };
  const apply = async (plan: unknown): Promise<string> => {
    const parsed: TemplatePlan = JSON.parse(JSON.stringify(plan));
    return codeOf(await applyTemplatePlan({ source, template, plan: parsed }));
  };
  const firstItem = (plan: Json): Json => {
    const body = plan["body"];
    const item: unknown = Array.isArray(body) ? body.at(0) : undefined;
    if (!isJson(item)) {
      throw new Error("the fixture plan has no first item");
    }
    return item;
  };
  const cases: [string, (plan: Record<string, unknown>) => unknown][] = [
    ["no plan at all", () => null],
    ["an array", () => []],
    ["a null body item", (plan) => ({ ...plan, body: [null] })],
    ["a number as body item", (plan) => ({ ...plan, body: [5] })],
    [
      "a source item without a block",
      (plan) => ({ ...plan, body: [{ kind: "source" }] }),
    ],
    [
      "an unknown item kind",
      (plan) => ({ ...plan, body: [{ kind: "other" }] }),
    ],
    [
      "a slot without source blocks",
      (plan) => ({
        ...plan,
        body: [{ kind: "slot", templateBlockId: "p0", sourceBlockIds: [] }],
      }),
    ],
    ["a null omit entry", (plan) => ({ ...plan, omit: [null] })],
    [
      "an unknown omit reason",
      (plan) => ({ ...plan, omit: [{ blockId: "p0", reason: "boring" }] }),
    ],
    [
      "references that are not a list",
      (plan) => ({ ...plan, references: "all" }),
    ],
    ["a null reference", (plan) => ({ ...plan, references: [null] })],
    ["a null highlight", (plan) => ({ ...plan, highlights: [null] })],
    [
      "numbering patches that are not a list",
      (plan) => ({ ...plan, numberingPatches: {} }),
    ],
    [
      "field choices that are not an object",
      (plan) => ({ ...plan, fields: 5 }),
    ],
    [
      "a field kind in lower case",
      (plan) => ({ ...plan, fields: { freezeKinds: ["date"] } }),
    ],
    [
      "a list that is neither an id nor new",
      (plan) => {
        firstItem(plan)["numbering"] = { numId: "x", ilvl: 0 };
        return plan;
      },
    ],
    [
      "a list level out of range",
      (plan) => {
        firstItem(plan)["numbering"] = { numId: 1, ilvl: 99 };
        return plan;
      },
    ],
    [
      "a negative restart",
      (plan) => {
        firstItem(plan)["numbering"] = { numId: 1, ilvl: 0, restartFrom: -1 };
        return plan;
      },
    ],
    [
      "nested placements that are not a list",
      (plan) => {
        firstItem(plan)["nested"] = "all";
        return plan;
      },
    ],
    [
      "a nested block placed twice",
      (plan) => {
        firstItem(plan)["nested"] = [{ blockId: "p9" }, { blockId: "p9" }];
        return plan;
      },
    ],
    [
      "an unknown merge",
      (plan) => {
        firstItem(plan)["mergeNext"] = "glue";
        return plan;
      },
    ],
    [
      "two patches of one level",
      (plan) => ({
        ...plan,
        numberingPatches: [
          { abstractNumId: 0, ilvl: 0, lvlText: "%1" },
          { abstractNumId: 0, ilvl: 0, start: 2 },
        ],
      }),
    ],
    [
      "an unknown number format",
      (plan) => ({
        ...plan,
        numberingPatches: [{ abstractNumId: 0, ilvl: 0, numFmt: "arabic" }],
      }),
    ],
    [
      "a blank face of a level",
      (plan) => ({
        ...plan,
        numberingPatches: [{ abstractNumId: 0, ilvl: 0, font: " " }],
      }),
    ],
    [
      "tab stops that do not ascend",
      (plan) => {
        firstItem(plan)["tabStops"] = [2552, 2552];
        return plan;
      },
    ],
    [
      "an empty list of tab stops",
      (plan) => {
        firstItem(plan)["tabStops"] = [];
        return plan;
      },
    ],
    [
      "an alignment that is none",
      (plan) => {
        firstItem(plan)["alignment"] = "middle";
        return plan;
      },
    ],
    [
      "a tab stop that is no whole number",
      (plan) => {
        firstItem(plan)["tabStops"] = [851.5];
        return plan;
      },
    ],
    [
      "a label break that is no boolean",
      (plan) => {
        firstItem(plan)["stripLabelBreak"] = "yes";
        return plan;
      },
    ],
    [
      "a negative start",
      (plan) => ({
        ...plan,
        numberingPatches: [{ abstractNumId: 0, ilvl: 0, start: -1 }],
      }),
    ],
    [
      "an indent that states nothing",
      (plan) => ({
        ...plan,
        numberingPatches: [{ abstractNumId: 0, ilvl: 0, indent: {} }],
      }),
    ],
    [
      "an indent in something other than whole twentieths of a point",
      (plan) => ({
        ...plan,
        numberingPatches: [
          { abstractNumId: 0, ilvl: 0, indent: { left: "1cm" } },
        ],
      }),
    ],
    [
      "an indent with an unknown side",
      (plan) => ({
        ...plan,
        numberingPatches: [{ abstractNumId: 0, ilvl: 0, indent: { top: 0 } }],
      }),
    ],
    [
      "a hanging and a first-line indent at once",
      (plan) => ({
        ...plan,
        numberingPatches: [
          { abstractNumId: 0, ilvl: 0, indent: { hanging: 0, firstLine: 0 } },
        ],
      }),
    ],
    [
      "cached text with a line end",
      (plan) => ({
        ...plan,
        references: [
          {
            blockId: "p0",
            start: 0,
            end: 5,
            targetBlockId: "p1",
            instruction: "\\h",
            expectedText: "a\nb",
          },
        ],
      }),
    ],
    [
      "cached text with a control character",
      (plan) => ({
        ...plan,
        references: [
          {
            blockId: "p0",
            start: 0,
            end: 5,
            targetBlockId: "p1",
            instruction: "\\h",
            expectedText: "a\u0001b",
          },
        ],
      }),
    ],
    [
      "cached text with half a character",
      (plan) => ({
        ...plan,
        references: [
          {
            blockId: "p0",
            start: 0,
            end: 5,
            targetBlockId: "p1",
            instruction: "\\h",
            expectedText: "\uD83D",
          },
        ],
      }),
    ],
    [
      "a switch with an argument",
      (plan) => ({
        ...plan,
        references: [
          {
            blockId: "p0",
            start: 0,
            end: 5,
            targetBlockId: "p1",
            instruction: "\\* Roman",
            expectedText: "First",
          },
        ],
      }),
    ],
    [
      "a switch in upper case",
      (plan) => ({
        ...plan,
        references: [
          {
            blockId: "p0",
            start: 0,
            end: 5,
            targetBlockId: "p1",
            instruction: "\\N",
            expectedText: "First",
          },
        ],
      }),
    ],
  ];

  test.each(cases)("%s is plan-invalid", async (_name, edit) => {
    expect(await apply(mutated(edit))).toBe("plan-invalid");
  });

  test("the valid plan the cases start from applies", async () => {
    expect(await apply(valid)).toBe("applied");
  });

  test("a joined item cannot carry a style of its own", async () => {
    const [first, second] = valid.body;
    if (first?.kind !== "source" || second?.kind !== "source") {
      throw new Error("fixture plan changed");
    }

    expect(
      await apply({
        ...valid,
        body: [
          { ...first, mergeNext: "space" },
          { ...second, styleId: "HouseBody" },
        ],
      }),
    ).toBe("merge-invalid");
  });

  test("a reference cannot target its own paragraph", async () => {
    expect(
      await apply({
        ...valid,
        references: [
          {
            blockId: "p0",
            start: 0,
            end: 5,
            targetBlockId: "p0",
            instruction: "\\h",
            expectedText: "First",
          },
        ],
      }),
    ).toBe("reference-span-invalid");
  });
});

describe("spans and characters", () => {
  test("a span may not cut a character in two", async () => {
    const source = await buildMinimalDocx(
      `${p(r("\u{1F4C4} see 1.1"))}${p(r("Target"))}`,
    );
    const reference = {
      blockId: "p0",
      targetBlockId: "p1",
      instruction: "\\h",
      expectedText: "x",
    };

    expect(
      codeOf(
        await applyFor({
          source,
          template,
          extra: {
            highlights: [{ blockId: "p0", start: 0, end: 1, color: "yellow" }],
          },
        }),
      ),
    ).toBe("highlight-span-invalid");
    expect(
      codeOf(
        await applyFor({
          source,
          template,
          extra: { references: [{ ...reference, start: 1, end: 5 }] },
        }),
      ),
    ).toBe("reference-span-invalid");
    // The whole character is a span like any other.
    expect(
      codeOf(
        await applyFor({
          source,
          template,
          extra: {
            highlights: [{ blockId: "p0", start: 0, end: 2, color: "yellow" }],
          },
        }),
      ),
    ).toBe("applied");
  });

  test("a span may not cut a symbol that takes two units", async () => {
    const source = await buildMinimalDocx(
      `${p(`<w:r><w:sym w:font="Segoe UI Emoji" w:char="1F600"/></w:r>${r(" smile")}`)}${p(r("Target"))}`,
    );

    expect(
      codeOf(
        await applyFor({
          source,
          template,
          extra: {
            references: [
              {
                blockId: "p0",
                start: 0,
                end: 1,
                targetBlockId: "p1",
                instruction: "\\h",
                expectedText: "x",
              },
            ],
          },
        }),
      ),
    ).toBe("reference-span-invalid");
  });

  test("a reference that changes text may not cover mixed emphasis", async () => {
    const body = (second: string): string =>
      `${p(`${r("see ")}${r("X", "<w:b/>")}${r("I", second)}`)}${p(r("Target"))}`;
    const references = [
      {
        blockId: "p0",
        start: 4,
        end: 6,
        targetBlockId: "p1",
        instruction: "\\h",
        expectedText: "12",
      },
    ];

    expect(
      codeOf(
        await applyFor({
          source: await buildMinimalDocx(body("")),
          template,
          extra: { references },
        }),
      ),
    ).toBe("reference-span-invalid");
    expect(
      codeOf(
        await applyFor({
          source: await buildMinimalDocx(body("<w:b/>")),
          template,
          extra: { references },
        }),
      ),
    ).toBe("applied");
  });
});

describe("label strips", () => {
  test.each([
    ["1.", true],
    ["1.1.", true],
    ["(a)", true],
    ["a)", true],
    ["Article I", true],
    ["Article 12", true],
    ["§ 5", true],
    ["XIII.", true],
    ["Annex No. 1", true],
    ["-", true],
    ["1. 1.", true],
    ["12. 3. 1.", true],
    ["1. 1", false],
    ["1. 1. 2026", false],
    ["", false],
    ["The", false],
    ["civil", false],
    ["Part One", false],
    ["1. The fee", false],
    ["An agreement between four parties and 1", false],
  ])("isManualLabel(%p) is %p", (label, expected) => {
    expect(isManualLabel(label)).toBe(expected);
  });

  const strip = async (text: string, stripLabel: string): Promise<string> =>
    codeOf(
      await applyFor({
        source: await buildMinimalDocx(p(r(text))),
        template,
        items: styled("HouseBody", { stripLabel }),
      }),
    );

  test("a label goes only when the block starts with it and a word ends there", async () => {
    expect(await strip("1. The fee is due.", "1.")).toBe("applied");
    // "1." is the start of a number here, not a label.
    expect(await strip("1.5 million shall be paid.", "1.")).toBe(
      "strip-label-mismatch",
    );
    expect(await strip("The parties agree.", "The")).toBe(
      "strip-label-mismatch",
    );
    expect(await strip("Articles follow.", "Article")).toBe(
      "strip-label-mismatch",
    );
  });
});

describe("emphasis from styles", () => {
  const STYLES = (defaults: string): string =>
    `${defaults}<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Struck"><w:name w:val="Struck"/><w:rPr><w:dstrike/></w:rPr></w:style><w:style w:type="character" w:styleId="StruckToo"><w:name w:val="Struck Too"/><w:rPr><w:dstrike/></w:rPr></w:style><w:style w:type="character" w:styleId="Strong"><w:name w:val="Strong"/><w:rPr><w:b/></w:rPr></w:style><w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:u w:val="single"/></w:rPr></w:style>`;

  test("double strike in both styles stays; it is not a toggle", async () => {
    const source = await buildSourceWith({
      body: p(
        r("struck twice", '<w:rStyle w:val="StruckToo"/>'),
        style("Struck"),
      ),
      styles: STYLES(""),
    });
    const { docx, report } = unwrap(await applyFor({ source, template }));

    expect(report.emphasisSpans).toMatchObject({ expected: 0, output: 0 });
    expect(await documentOf(docx)).not.toContain("dstrike");
  });

  test("a character style inside a link is emphasis unless it is the link's look", async () => {
    const source = await buildSourceWith({
      body: p(
        `<w:hyperlink r:id="rId20">${r("bold link", '<w:rStyle w:val="Strong"/>')}${r(" plain link", '<w:rStyle w:val="Hyperlink"/>')}</w:hyperlink>`,
      ),
      styles: STYLES(""),
      relationships: [
        {
          id: "rId20",
          type: "hyperlink",
          target: "https://example.org/",
          external: true,
        },
      ],
    });
    const { docx, report } = unwrap(await applyFor({ source, template }));
    const document = await documentOf(docx);

    expect(report.emphasisSpans).toMatchObject({ expected: 1, output: 1 });
    expect(document).toContain(
      '<w:r><w:rPr><w:rStyle w:val="Hyperlink"/><w:b/></w:rPr><w:t xml:space="preserve">bold link</w:t></w:r><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t xml:space="preserve"> plain link</w:t></w:r>',
    );
  });

  test("document defaults are the baseline a run departs from", async () => {
    const source = await buildSourceWith({
      body: p(`${r("bold by default")}${r(" not bold", '<w:b w:val="0"/>')}`),
      styles: STYLES(
        "<w:docDefaults><w:rPrDefault><w:rPr><w:b/></w:rPr></w:rPrDefault></w:docDefaults>",
      ),
    });
    const { docx } = unwrap(await applyFor({ source, template }));

    expect(await documentOf(docx)).toContain(
      '<w:r><w:t xml:space="preserve">bold by default</w:t></w:r><w:r><w:rPr><w:b w:val="0"/></w:rPr><w:t xml:space="preserve"> not bold</w:t></w:r>',
    );
  });
});

describe("links", () => {
  const external = {
    id: "rId20",
    type: "hyperlink",
    target: "https://example.org/doc",
    external: true,
  };

  test("the place a link names inside its address is carried over", async () => {
    const source = await buildSourceWith({
      body: p(
        `<w:hyperlink r:id="rId20" w:anchor="section2">${r("section two")}</w:hyperlink>`,
      ),
      relationships: [external],
    });
    const { docx, report } = unwrap(await applyFor({ source, template }));

    expect(report.ok).toBe(true);
    expect(await documentOf(docx)).toMatch(
      /<w:hyperlink r:id="rId\d+" w:anchor="section2" w:history="1">/u,
    );
  });

  test("a HYPERLINK field keeps its address and place", async () => {
    const source = await buildMinimalDocx(
      p(
        fieldRuns(
          'HYPERLINK "https://example.org/doc" \\l "part2"',
          r("part two"),
        ),
      ),
    );
    const { docx, report } = unwrap(await applyFor({ source, template }));

    expect(report.dropped.sourceFields).toEqual([]);
    expect(await documentOf(docx)).toMatch(
      /<w:hyperlink r:id="rId\d+" w:anchor="part2" w:history="1">/u,
    );
    expect(await readPart(docx, "word/_rels/document.xml.rels")).toContain(
      'Target="https://example.org/doc" TargetMode="External"',
    );
  });

  test("a link into the source package becomes text and is counted", async () => {
    const source = await buildSourceWith({
      body: p(
        `${r("See ")}<w:hyperlink r:id="rId20">${r("the other part")}</w:hyperlink>`,
      ),
      relationships: [{ ...external, target: "other.xml", external: false }],
    });
    const { docx, report } = unwrap(await applyFor({ source, template }));

    expect(report.ok).toBe(true);
    expect(report.dropped.internalHyperlinks).toBe(1);
    expect(await documentOf(docx)).not.toContain("<w:hyperlink r:id");
  });

  test("verification notices a link that lost its wrapper", async () => {
    const source = await buildSourceWith({
      body: p(`<w:hyperlink r:id="rId20">${r("linked")}</w:hyperlink>`),
      relationships: [external],
    });
    const plan = await planFor({ source, template });
    const { docx } = unwrap(
      await applyTemplatePlan({ source, template, plan }),
    );
    const document = await documentOf(docx);
    const unlinked = document
      .replace(/<w:hyperlink r:id="rId\d+" w:history="1">/u, "")
      .replace("</w:hyperlink>", "");
    expect(unlinked).not.toBe(document);
    const report = await verifyFor({
      source,
      template,
      output: await replacePart(docx, "word/document.xml", unlinked),
      plan,
    });

    expect(report.issues.map(({ code }) => code)).toContain("text-mismatch");
  });
});

describe("list instances and patches", () => {
  const withNumbering = async (
    edit: (xml: string) => string,
  ): Promise<Uint8Array> => {
    const xml = (await readPart(template, "word/numbering.xml")) ?? "";
    const edited = edit(xml);
    expect(edited).not.toBe(xml);
    return await replacePart(template, "word/numbering.xml", edited);
  };
  const OVERRIDE =
    '<w:lvlOverride w:ilvl="1"><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="upperLetter"/><w:lvlText w:val="%2)"/><w:lvlJc w:val="left"/></w:lvl></w:lvlOverride>';
  const three = async (): Promise<Uint8Array> =>
    await buildMinimalDocx(`${p(r("one"))}${p(r("two"))}${p(r("three"))}`);
  const grouped =
    (levels: { ilvl: number; restartFrom?: number }[]) =>
    (blocks: PlannedBlock[]): TemplatePlanItem[] =>
      blocks.map(({ id, hash }, index) => ({
        kind: "source",
        blockId: id,
        hash,
        styleId: "HouseBody",
        numbering: {
          numId: "new",
          baseNumId: 1,
          group: "g",
          ...(levels[index] ?? { ilvl: 0 }),
        },
      }));

  test("a new instance keeps its base list's overrides and restarts each named level", async () => {
    const overridden = await withNumbering((xml) =>
      xml.replace(
        '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>',
        `<w:num w:numId="1"><w:abstractNumId w:val="0"/>${OVERRIDE}</w:num>`,
      ),
    );
    const { docx } = unwrap(
      await applyFor({
        source: await three(),
        template: overridden,
        items: grouped([{ ilvl: 0 }, { ilvl: 1, restartFrom: 5 }, { ilvl: 1 }]),
      }),
    );

    expect(await readPart(docx, "word/numbering.xml")).toContain(
      '<w:num w:numId="3"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride><w:lvlOverride w:ilvl="1"><w:startOverride w:val="5"/><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="upperLetter"/><w:lvlText w:val="%2)"/><w:lvlJc w:val="left"/></w:lvl></w:lvlOverride></w:num>',
    );
  });

  test("one level of a group cannot restart from two numbers", async () => {
    expect(
      codeOf(
        await applyFor({
          source: await three(),
          template,
          items: grouped([
            { ilvl: 0 },
            { ilvl: 1, restartFrom: 5 },
            { ilvl: 1, restartFrom: 9 },
          ]),
        }),
      ),
    ).toBe("plan-invalid");
  });

  test("a patched start is where a new instance starts", async () => {
    const { docx } = unwrap(
      await applyFor({
        source: await three(),
        template,
        items: grouped([{ ilvl: 0 }, { ilvl: 0 }, { ilvl: 0 }]),
        extra: {
          numberingPatches: [{ abstractNumId: 0, ilvl: 0, start: 4 }],
        },
      }),
    );

    expect(await readPart(docx, "word/numbering.xml")).toContain(
      '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="4"/></w:lvlOverride>',
    );
    expect(
      unwrap(await analyzeDocxBlocks(docx)).blocks.map(({ label }) => label),
    ).toEqual(["Art. 4", "Art. 5", "Art. 6"]);
  });

  test("a group named like an ungrouped item's key stays its own list", async () => {
    const { docx } = unwrap(
      await applyFor({
        source: await three(),
        template,
        items: (blocks) =>
          blocks.map(({ id, hash }, index) => ({
            kind: "source",
            blockId: id,
            hash,
            styleId: "HouseBody",
            numbering: {
              numId: "new",
              baseNumId: 1,
              ilvl: 0,
              ...(index === 1 ? { group: "#1" } : {}),
            },
          })),
      }),
    );

    expect(
      ((await readPart(docx, "word/numbering.xml")) ?? "").match(
        /<w:num w:numId="[345]">/gu,
      ),
    ).toHaveLength(3);
  });

  test("a level that wraps its format in a compatibility alternative is not patched", async () => {
    const alternate = await withNumbering((xml) =>
      xml.replace(
        /(<w:lvl w:ilvl="4"><w:start w:val="1"\/>)<w:numFmt w:val="lowerRoman"\/>/u,
        '$1<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="w14"><w:numFmt w:val="custom" w:format="001, 002"/></mc:Choice><mc:Fallback><w:numFmt w:val="decimal"/></mc:Fallback></mc:AlternateContent>',
      ),
    );

    expect(
      codeOf(
        await applyFor({
          source: await three(),
          template: alternate,
          extra: {
            numberingPatches: [
              { abstractNumId: 0, ilvl: 4, numFmt: "decimal" },
            ],
          },
        }),
      ),
    ).toBe("unsupported-content");
  });
});

describe("template fields", () => {
  const nestedFooter = (inner: string): string =>
    `${XML}<w:ftr ${NS}>${p(`<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> IF </w:instrText></w:r>${inner}<w:r><w:instrText> = "x" "a" "b" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${r("b")}<w:r><w:fldChar w:fldCharType="end"/></w:r>`)}</w:ftr>`;

  test("a field of a frozen kind inside another field's code stays, with a warning", async () => {
    const nested = await replacePart(
      template,
      "word/footer1.xml",
      nestedFooter(fieldRuns("DATE", r("1.1.2020"))),
    );
    const { docx, report } = unwrap(
      await applyFor({
        source: await buildMinimalDocx(p(r("Body text."))),
        template: nested,
        extra: { fields: { freezeKinds: ["DATE"] } },
      }),
    );

    expect(report.ok).toBe(true);
    expect(report.warnings.join(" ")).toContain("inside another field's code");
    expect(await readPart(docx, "word/footer1.xml")).toContain("DATE");
  });

  test("a prompting field in another field's code still fails under update on open", async () => {
    const nested = await replacePart(
      template,
      "word/footer1.xml",
      nestedFooter(fieldRuns('FILLIN "x"', r("y"))),
    );
    const applied = await applyFor({
      source: await buildMinimalDocx(p(r("Body text."))),
      template: nested,
      extra: { fields: { freezeKinds: ["FILLIN"], updateOnOpen: true } },
    });

    expect(codeOf(applied)).toBe("verification-failed");
    expect(issuesOf(applied)).toEqual(["prompting-field-left"]);
  });

  test("verification notices source text inside a field of the template", async () => {
    const source = await buildMinimalDocx(
      `${p(r("First clause."))}${p(r("Second clause."))}`,
    );
    const plan = await planFor({
      source,
      template,
      items: (blocks) => [
        { kind: "template", templateBlockIds: ["p4", "p5"] },
        ...blocks.map(({ id, hash }): TemplatePlanItem => ({
          kind: "source",
          blockId: id,
          hash,
        })),
      ],
    });
    const { docx } = unwrap(
      await applyTemplatePlan({ source, template, plan }),
    );
    const document = await documentOf(docx);
    // Move the end of the contents field behind the last source paragraph.
    const END = '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
    const close = document.indexOf(
      `${END}</w:p>`,
      document.indexOf("Second sample"),
    );
    const moved =
      `${document.slice(0, close)}${document.slice(close + END.length)}`.replace(
        "Second clause.</w:t></w:r>",
        `Second clause.</w:t></w:r>${END}`,
      );
    expect(moved).not.toBe(document);
    const report = await verifyFor({
      source,
      template,
      output: await replacePart(docx, "word/document.xml", moved),
      plan,
    });

    expect(report.ok).toBe(false);
    expect(report.issues.map(({ code }) => code)).toContain("field-unbalanced");
  });
});

describe("a kept table of contents", () => {
  const TOC_BEGIN =
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC';
  const FIELD_END = '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
  const EMPTIED = `<w:r><w:fldChar w:fldCharType="separate"/></w:r>${FIELD_END}</w:p>`;

  /** The template with its document part rewritten. */
  const reshaped = async (
    change: (document: string) => string,
  ): Promise<Uint8Array> => {
    const document = await documentOf(template);
    const changed = change(document);
    expect(changed).not.toBe(document);
    return await replacePart(template, "word/document.xml", changed);
  };
  const withContents = (blocks: PlannedBlock[]): TemplatePlanItem[] => [
    {
      kind: "template",
      templateBlockIds: [TEMPLATE_IDS.tocFirst, TEMPLATE_IDS.tocLast],
    },
    ...blocks.map(({ id, hash }): TemplatePlanItem => ({
      kind: "source",
      blockId: id,
      hash,
    })),
  ];
  const applyOnto = async (
    onto: Uint8Array,
    extra: Partial<TemplatePlan> = {},
  ) => {
    const source = await buildMinimalDocx(
      `${p(r("First clause."))}${p(r("Second clause."))}`,
    );
    const plan = await planFor({
      source,
      template: onto,
      items: withContents,
      extra,
    });
    const applied = unwrap(
      await applyTemplatePlan({ source, template: onto, plan }),
    );
    return {
      ...applied,
      source,
      plan,
      document: await documentOf(applied.docx),
    };
  };

  test("text and a whole field in front of the contents field stay; the contents field is the one emptied", async () => {
    const page = fieldRuns("PAGE", r("1"));
    const onto = await reshaped((document) =>
      document.replace(TOC_BEGIN, `${r("Page ")}${page}${r(": ")}${TOC_BEGIN}`),
    );
    const { document, report } = await applyOnto(onto);

    expect(report.ok).toBe(true);
    expect(document).toContain(`${r("Page ")}${page}${r(": ")}${TOC_BEGIN}`);
    expect(document).toContain(`\\z </w:instrText></w:r>${EMPTIED}`);
    expect(document).not.toContain("Sample article");
    expect(document).not.toContain("Second sample");
  });

  test("text behind the field's own end keeps the contents as the template has them, whatever field follows it", async () => {
    const onto = await reshaped((document) => {
      const close = document.indexOf(
        `${FIELD_END}</w:p>`,
        document.indexOf("Second sample"),
      );
      return `${document.slice(0, close + FIELD_END.length)}${r(" Printed on page ")}${fieldRuns("PAGE", r("3"))}${document.slice(close + FIELD_END.length)}`;
    });
    const { document } = await applyOnto(onto);

    expect(document).toContain("Sample article");
    expect(document).toContain("Second sample");
    expect(document).toContain(" Printed on page ");
  });

  test("a bookmark that starts in front of the field and ends among its entries goes with them", async () => {
    // Both halves in the paragraph the field begins in: the end among the
    // entries the paragraph loses, the start in front of the field.
    const onto = await reshaped((document) => {
      const close = document.indexOf(
        "</w:p>",
        document.indexOf("Sample article"),
      );
      return `${document.slice(0, close)}<w:bookmarkEnd w:id="77"/>${document.slice(close)}`.replace(
        TOC_BEGIN,
        `<w:bookmarkStart w:id="77" w:name="ContentsMark"/>${TOC_BEGIN}`,
      );
    });
    const { document, report } = await applyOnto(onto);

    expect(report.ok).toBe(true);
    expect(document).toContain(EMPTIED);
    expect(document).not.toContain("ContentsMark");
    expect(document).not.toContain('<w:bookmarkEnd w:id="77"/>');
  });

  test("a plan that freezes the field keeps its entries as text and adds no end of a field", async () => {
    const { document, report } = await applyOnto(template, {
      fields: { freezeKinds: ["TOC"] },
    });

    expect(report.ok).toBe(true);
    expect(document).not.toContain(" TOC ");
    expect(document).toContain("Sample article");
    expect(document).toContain("Second sample");
    expect(document.split('w:fldCharType="begin"')).toHaveLength(
      document.split('w:fldCharType="end"').length,
    );
  });

  test("verification holds the emptied paragraph to the template's style", async () => {
    const { docx, document, source, plan } = await applyOnto(template);
    const emptied = document.indexOf(TOC_BEGIN);
    const paragraph = document.lastIndexOf("<w:p>", emptied);
    const restyled = `${document.slice(0, paragraph)}${document
      .slice(paragraph, emptied)
      .replace('w:val="TOC1"', 'w:val="HouseBody"')}${document.slice(emptied)}`;
    expect(restyled).not.toBe(document);
    const report = await verifyFor({
      source,
      template,
      output: await replacePart(docx, "word/document.xml", restyled),
      plan,
    });

    expect(report.ok).toBe(false);
    expect(report.issues.map(({ code }) => code)).toContain(
      "template-mismatch",
    );
  });
});

describe("a slot that drops the name of a field the template shows", () => {
  // The template's table has two cells: the left one stands for the cell
  // that shows the field, the right one for the cell that takes its value.
  const shower = TEMPLATE_IDS.signatureLeft;
  const cell = TEMPLATE_IDS.signatureRight;
  const planOf = async (
    line: string,
    dropLead: { text: string; shownByTemplateBlockId: string },
    extra: Partial<TemplatePlan> = {},
  ) => {
    const source = await buildMinimalDocx(
      `${p(`${r(line.split("\t").at(0) ?? "")}${line.includes("\t") ? `<w:r><w:tab/></w:r>${r(line.split("\t").at(1) ?? "")}` : ""}`)}${p(r("The body of the note."))}`,
    );
    const plan = await planFor({
      source,
      template,
      items: ([head, ...rest]) => [
        { kind: "template", templateBlockIds: [TEMPLATE_IDS.signatures] },
        {
          kind: "slot",
          templateBlockId: cell,
          sourceBlockIds: head === undefined ? [] : [head.id],
          dropLead,
        },
        ...rest.map(({ id, hash }): TemplatePlanItem => ({
          kind: "source",
          blockId: id,
          hash,
        })),
      ],
      extra,
    });
    return { source, plan };
  };

  test("the lead goes, the value stands word for word, and the report names the lead", async () => {
    const { source, plan } = await planOf("To:\tJane Roe, Legal", {
      text: "To:",
      shownByTemplateBlockId: shower,
    });
    const { docx, report } = unwrap(
      await applyTemplatePlan({ source, template, plan }),
    );
    const document = await documentOf(docx);

    expect(report.ok).toBe(true);
    expect(document).toContain(
      '<w:r><w:t xml:space="preserve">Jane Roe, Legal</w:t></w:r></w:p></w:tc>',
    );
    expect(document).not.toContain("To:");
    expect(report.blocks.find(({ blockId }) => blockId === "p0")).toMatchObject(
      {
        disposition: "slot",
        text: "identical",
        droppedLead: { text: "To:\t", shownByTemplateBlockId: shower },
      },
    );
    // No label of the source is reported as lost with it.
    expect(
      report.blocks.find(({ blockId }) => blockId === "p0")?.sourceLabel,
    ).toBeNull();
    // Verification holds the value to the source: the lead back in the
    // cell, or another word in the value, is no longer what the plan placed.
    for (const changed of [
      document.replace(">Jane Roe, Legal<", ">To: Jane Roe, Legal<"),
      document.replace(">Jane Roe, Legal<", ">Jane Doe, Legal<"),
    ]) {
      expect(changed).not.toBe(document);
      const checked = await verifyFor({
        source,
        template,
        output: await replacePart(docx, "word/document.xml", changed),
        plan,
      });
      expect(checked.ok).toBe(false);
    }
  });

  test("a line that does not start with the lead, or holds nothing behind it, is refused", async () => {
    for (const [line, lead] of [
      ["Attention:\tJane Roe", "To:"],
      ["To:", "To:"],
      ["To:\tJane Roe", " "],
    ] as const) {
      const { source, plan } = await planOf(line, {
        text: lead,
        shownByTemplateBlockId: shower,
      });
      const applied = await applyTemplatePlan({ source, template, plan });
      expect(codeOf(applied)).toBe("strip-label-mismatch");
    }
  });

  test("the paragraph that shows the field must be kept as the template has it, with text", async () => {
    const refused = async (
      shownByTemplateBlockId: string,
      extra: Partial<TemplatePlan> = {},
    ) => {
      const { source, plan } = await planOf(
        "To:\tJane Roe",
        { text: "To:", shownByTemplateBlockId },
        extra,
      );
      return codeOf(await applyTemplatePlan({ source, template, plan }));
    };
    // A paragraph the plan does not keep, one that is no block of the
    // template, the slot itself, and one the plan empties.
    expect(await refused(TEMPLATE_IDS.party)).toBe("unknown-block");
    expect(await refused("p999")).toBe("unknown-block");
    expect(await refused(cell)).toBe("unknown-block");
    expect(await refused(shower, { emptyTemplateBlockIds: [shower] })).toBe(
      "unknown-block",
    );
  });
});

describe("drawn shapes fixed to the page", () => {
  const WP =
    "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
  const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
  const WPS =
    "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
  const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
  /** A floating shape: a dot, or a text box when it has text of its own. */
  const shape = (id: number, horizontal: string, inner = ""): string =>
    `<w:r><mc:AlternateContent xmlns:mc="${MC}"><mc:Choice xmlns:wps="${WPS}" Requires="wps"><w:drawing><wp:anchor xmlns:wp="${WP}" distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="${String(id)}" behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="${horizontal}"><wp:posOffset>967105</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>1197610</wp:posOffset></wp:positionV><wp:extent cx="59055" cy="59055"/><wp:wrapNone/><wp:docPr id="${String(id)}" name="Shape ${String(id)}"/><a:graphic xmlns:a="${A}"><a:graphicData uri="${WPS}"><wps:wsp><wps:spPr/>${inner}<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice><mc:Fallback><w:pict/></mc:Fallback></mc:AlternateContent></w:r>`;
  const BODY = [
    // Two marks of the page, one shape that follows its column, and a text
    // box that is fixed to the page but says something.
    p(
      `${shape(1, "page")}${shape(2, "margin")}${shape(3, "column")}${shape(4, "page", `<wps:txbx><w:txbxContent>${p(r("Note"))}</w:txbxContent></wps:txbx>`)}${r("Text of the paragraph.")}`,
    ),
    p(r("Plain paragraph.")),
  ].join("");
  const planOf = async (source: Uint8Array, drop: boolean | undefined) =>
    await planFor({
      source,
      template,
      items: (blocks) =>
        blocks.map(({ id, hash }): TemplatePlanItem => {
          const item: TemplatePlanItem = {
            kind: "source",
            blockId: id,
            hash,
            styleId: "HouseBody",
          };
          if (id === "p0" && drop !== undefined) {
            item.dropPageShapes = drop;
          }
          return item;
        }),
    });

  test("the manifest counts them apart from a paragraph's other objects", async () => {
    const [first, second] = unwrap(
      await analyzeDocxBlocks(await buildMinimalDocx(BODY)),
    ).blocks;
    expect(first).toMatchObject({ inlineObjects: 4, pageShapes: 2 });
    expect(second?.pageShapes).toBeUndefined();
  });

  test("a paragraph is written without them when its item says so, and with everything else", async () => {
    const source = await buildMinimalDocx(BODY);
    const dropping = await planOf(source, true);
    const keeping = await planOf(source, undefined);
    const dropped = unwrap(
      await applyTemplatePlan({ source, template, plan: dropping }),
    );
    const kept = unwrap(
      await applyTemplatePlan({ source, template, plan: keeping }),
    );
    const anchorsOf = async (docx: Uint8Array) =>
      [
        ...(await documentOf(docx)).matchAll(
          /relativeFrom="(\w+)"><wp:posOffset>967105/gu,
        ),
      ].map(([, from]) => from);

    expect(dropped.report.ok).toBe(true);
    expect(await anchorsOf(dropped.docx)).toEqual(["column", "page"]);
    expect(await documentOf(dropped.docx)).toContain("Text of the paragraph.");
    expect(dropped.report.dropped.pageShapes).toBe(2);
    // Without the option every object is carried.
    expect(await anchorsOf(kept.docx)).toEqual([
      "page",
      "margin",
      "column",
      "page",
    ]);
    expect(kept.report.dropped.pageShapes).toBe(0);
    // Verification holds the output to what the plan names, both ways.
    for (const [output, plan] of [
      [dropped.docx, keeping],
      [kept.docx, dropping],
    ] as const) {
      expect((await verifyFor({ source, template, output, plan })).ok).toBe(
        false,
      );
    }
  });

  test("the option is a switch of a paragraph", async () => {
    const source = await buildMinimalDocx(
      `${BODY}<w:tbl><w:tr><w:tc>${p(r("cell"))}</w:tc></w:tr></w:tbl>${p(r("end"))}`,
    );
    const plan = await planFor({ source, template });
    const withOption = (blockId: string, value: unknown): TemplatePlan => {
      const body: TemplatePlanItem[] = [];
      for (const item of plan.body) {
        const changed = structuredClone(item);
        if (changed.kind === "source" && changed.blockId === blockId) {
          Object.assign(changed, { dropPageShapes: value });
        }
        body.push(changed);
      }
      return { ...plan, body };
    };
    expect(
      codeOf(
        await applyTemplatePlan({
          source,
          template,
          plan: withOption("p0", "yes"),
        }),
      ),
    ).toBe("plan-invalid");
    expect(
      codeOf(
        await applyTemplatePlan({
          source,
          template,
          plan: withOption("t2", true),
        }),
      ),
    ).toBe("plan-invalid");
  });
});

describe("labels written as text", () => {
  const level = (
    ilvl: number,
    numFmt: string,
    lvlText: string,
    suff = "",
  ): string =>
    `<w:lvl w:ilvl="${ilvl}"><w:start w:val="1"/><w:numFmt w:val="${numFmt}"/>${suff}<w:lvlText w:val="${lvlText}"/></w:lvl>`;
  const numbered = (text: string, numId: number, ilvl: number): string =>
    p(
      r(text),
      `<w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr>`,
    );
  const NUMBERING = [
    `<w:abstractNum w:abstractNumId="0">${level(0, "upperRoman", "%1.")}${level(1, "decimal", "%1.%2", '<w:suff w:val="space"/>')}${level(2, "ordinal", "%1.%3")}</w:abstractNum>`,
    `<w:abstractNum w:abstractNumId="1">${level(0, "bullet", "\u2022")}</w:abstractNum>`,
    '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>',
    '<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>',
  ].join("");
  const source = async (): Promise<Uint8Array> =>
    await buildSourceWith({
      body: [
        numbered("Scope", 1, 0),
        numbered("The works.", 1, 1),
        numbered("A point.", 2, 0),
        numbered("A spelled number.", 1, 2),
      ].join(""),
      numbering: NUMBERING,
    });
  const asText =
    (ids: readonly string[]) =>
    (blocks: PlannedBlock[]): TemplatePlanItem[] =>
      blocks.map(({ id, hash }) => ({
        kind: "source",
        blockId: id,
        hash,
        numbering: null,
        ...(ids.includes(id) ? { labelAsText: true } : {}),
      }));

  test("a number keeps the tab or space its list puts after it", async () => {
    const { docx, report } = unwrap(
      await applyFor({
        source: await source(),
        template,
        items: asText(["p0", "p1"]),
      }),
    );
    const blocks = unwrap(await analyzeDocxBlocks(docx)).blocks;
    const textOf = (blockId: string): string | undefined => {
      const outputBlockId = report.blocks.find(
        (block) => block.blockId === blockId,
      )?.outputBlockId;
      return blocks.find(({ id }) => id === outputBlockId)?.text;
    };

    expect(report.ok).toBe(true);
    expect(textOf("p0")).toBe("I.\tScope");
    expect(textOf("p1")).toBe("I.1 The works.");
    // The other two say so: their labels are gone.
    expect(report.counts).toMatchObject({ labelAsText: 2, labelLost: 2 });
  });

  test("a bullet and a number a language spells are not text", async () => {
    for (const blockId of ["p2", "p3"]) {
      const applied = await applyFor({
        source: await source(),
        template,
        items: asText([blockId]),
      });
      expect(codeOf(applied)).toBe("label-not-writable");
      if (Result.isError(applied)) {
        expect(applied.error.blockId).toBe(blockId);
      }
    }
  });
});

describe("number-only field predictions", () => {
  test("keep the punctuation after the counter and drop the blanks", async () => {
    const text = "See part A- above.";
    const { report } = unwrap(
      await applyFor({
        source: await buildSourceWith({
          body: [p(r("Heading")), p(r(text))].join(""),
        }),
        template,
        items: (blocks) =>
          blocks.map(({ id, hash }, index) => ({
            kind: "source",
            blockId: id,
            hash,
            numbering: index === 0 ? { numId: 1, ilvl: 0 } : null,
          })),
        extra: {
          numberingPatches: [
            {
              abstractNumId: 0,
              ilvl: 0,
              numFmt: "upperLetter",
              lvlText: "Part %1 -",
            },
          ],
          references: [
            {
              blockId: "p1",
              start: text.indexOf("A-"),
              end: text.indexOf("A-") + 2,
              targetBlockId: "p0",
              instruction: "\\n \\t \\h",
              expectedText: "A-",
            },
          ],
        },
      }),
    );

    // The words of `Part %1 -` go, and with them the blank before the dash.
    expect(report.references.results.at(0)).toMatchObject({
      status: "ok",
      predictedText: "A-",
    });
    expect(report.references.predictionMismatch).toBe(0);
  });
});

describe("reference fields with a letter-case format", () => {
  const cited = async (
    text: string,
    cite: string,
    instruction: string,
    lvlText = "Art. %1",
    numFmt = "decimal",
  ) =>
    await applyFor({
      source: await buildSourceWith({
        body: [p(r("Heading")), p(r(text))].join(""),
      }),
      template,
      items: (blocks) =>
        blocks.map(({ id, hash }, index) => ({
          kind: "source",
          blockId: id,
          hash,
          numbering: index === 0 ? { numId: 1, ilvl: 0 } : null,
        })),
      extra: {
        numberingPatches: [{ abstractNumId: 0, ilvl: 0, numFmt, lvlText }],
        references: [
          {
            blockId: "p1",
            start: text.indexOf(cite),
            end: text.indexOf(cite) + cite.length,
            targetBlockId: "p0",
            instruction,
            expectedText: cite,
          },
        ],
      },
    });

  test("a field over a citation's word and number prints the label in the citation's case", async () => {
    for (const [text, cite, instruction] of [
      ["See art. 1 above.", "art. 1", "\\n \\* lower \\h"],
      ["Art. 1 says so.", "Art. 1", "\\n \\h"],
      ["SEE ART. 1 ABOVE.", "ART. 1", "\\n \\* Upper \\h"],
    ] as const) {
      const { docx, report } = unwrap(await cited(text, cite, instruction));
      expect(report.references.results.at(0)).toMatchObject({
        status: "ok",
        predictedText: cite,
      });
      expect(report.references.predictionMismatch).toBe(0);
      expect(await readPart(docx, "word/document.xml")).toContain(
        ` ${instruction} </w:instrText>`,
      );
    }
  });

  test("a lower-case format lowers a roman numeral too, and the prediction says so", async () => {
    const { report } = unwrap(
      await cited(
        "See art. I above.",
        "art. I",
        "\\n \\* lower \\h",
        "Art. %1",
        "upperRoman",
      ),
    );
    expect(report.references.results.at(0)?.predictedText).toBe("art. i");
    expect(report.references.predictionMismatch).toBe(1);
  });

  test("refuses any other format switch", async () => {
    for (const instruction of [
      "\\n \\* Roman \\h",
      "\\n \\* MERGEFORMAT \\* CHARFORMAT \\h",
      "\\n \\* lower \\* Upper \\h",
      "\\n \\* \\h",
    ]) {
      expect(
        codeOf(await cited("See art. 1 above.", "art. 1", instruction)),
      ).toBe("plan-invalid");
    }
  });
});

describe("ordinal labels by the paragraph's language", () => {
  const level = (ilvl: number, numFmt: string, lvlText: string): string =>
    `<w:lvl w:ilvl="${ilvl}"><w:start w:val="1"/><w:numFmt w:val="${numFmt}"/><w:lvlText w:val="${lvlText}"/></w:lvl>`;
  const NUMBERING = [
    `<w:abstractNum w:abstractNumId="0">${level(0, "decimal", "%1.")}${level(1, "ordinal", "%1.%2")}</w:abstractNum>`,
    `<w:abstractNum w:abstractNumId="1">${level(0, "ordinal", "%1")}</w:abstractNum>`,
    '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>',
    '<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>',
  ].join("");
  const item = (
    text: string,
    numId: number,
    ilvl: number,
    extra = "",
  ): string =>
    p(
      r(text),
      `${extra}<w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr>`,
    );
  const marked = (language: string): string =>
    `<w:rPr><w:lang w:val="${language}"/></w:rPr>`;
  // The mark's language follows `w:numPr` in the schema; the reader takes either order.
  const source = async (styles?: string): Promise<Uint8Array> =>
    await buildSourceWith({
      body: [
        item("Article", 1, 0, marked("cs-CZ")),
        item("Clause", 1, 1, marked("cs-CZ")),
        item("First", 2, 0, marked("de-DE")),
        item("Second", 2, 0, marked("fr-FR")),
        item("Third", 2, 0, marked("en-GB")),
        item("Fourth", 2, 0, marked("xx-XX")),
        item("Fifth", 2, 0),
      ].join(""),
      numbering: NUMBERING,
      ...(styles === undefined ? {} : { styles }),
    });
  const labelsOf = async (docx: Uint8Array): Promise<(string | null)[]> =>
    unwrap(await analyzeDocxBlocks(docx)).blocks.map(({ label }) => label);

  test("a Czech clause numbered with the ordinal format reads 1.1.", async () => {
    expect(await labelsOf(await source())).toEqual([
      "1.",
      "1.1.",
      "1.",
      "2e",
      "3rd",
      // A language whose ordinals are not known, and no language at all.
      "4th",
      "5th",
    ]);
  });

  test("the language comes from the style and the document defaults too", async () => {
    const styles =
      '<w:docDefaults><w:rPrDefault><w:rPr><w:lang w:val="cs-CZ"/></w:rPr></w:rPrDefault></w:docDefaults>';
    expect((await labelsOf(await source(styles))).at(-1)).toBe("5.");
  });

  test("such a label can be written as text, one in an unknown language cannot", async () => {
    const asText =
      (id: string) =>
      (blocks: PlannedBlock[]): TemplatePlanItem[] =>
        blocks.map(({ id: blockId, hash }) => ({
          kind: "source",
          blockId,
          hash,
          numbering: null,
          ...(blockId === id ? { labelAsText: true } : {}),
        }));
    const czech = unwrap(
      await applyFor({
        source: await source(),
        template,
        items: asText("p1"),
      }),
    );
    expect(
      czech.report.blocks.find(({ blockId }) => blockId === "p1")?.labelAsText,
    ).toBe("1.1.");
    expect(
      codeOf(
        await applyFor({
          source: await source(),
          template,
          items: asText("p5"),
        }),
      ),
    ).toBe("label-not-writable");
  });
});

describe("a heading run into its text", () => {
  const HIDDEN_MARK = "<w:rPr><w:vanish/><w:specVanish/></w:rPr>";
  const source = async (): Promise<Uint8Array> =>
    await buildSourceWith({
      body: [
        p(r("Purchase and Sale"), HIDDEN_MARK),
        p(r(".   On the Closing Date the Company sells the Shares.")),
        p(r("Closing"), '<w:rPr><w:specVanish w:val="0"/></w:rPr>'),
      ].join(""),
    });

  test("the manifest says which paragraph runs on into the next", async () => {
    const { blocks } = unwrap(await analyzeDocxBlocks(await source()));
    expect(blocks.map(({ runsIn }) => runsIn ?? false)).toEqual([
      true,
      false,
      false,
    ]);
  });

  test("joined to its text, the heading leaves one paragraph whose mark shows", async () => {
    const applied = unwrap(
      await applyFor({
        source: await source(),
        template,
        items: (blocks) =>
          blocks.map(({ id, hash }, position): TemplatePlanItem =>
            position === 0
              ? { kind: "source", blockId: id, hash, mergeNext: "none" }
              : { kind: "source", blockId: id, hash },
          ),
      }),
    );
    expect(applied.report.ok).toBe(true);
    const xml = await documentOf(applied.docx);
    expect(xml).toContain("Purchase and Sale");
    expect(xml).not.toContain("specVanish");
    const { blocks } = unwrap(await analyzeDocxBlocks(applied.docx));
    expect(blocks.map(({ text }) => text)).toContain(
      "Purchase and Sale.   On the Closing Date the Company sells the Shares.",
    );
  });
});

describe("italic a style gives its text", () => {
  const STYLES = [
    '<w:style w:type="paragraph" w:styleId="NoteText"><w:name w:val="Note Text"/><w:rPr><w:i/><w:iCs/></w:rPr></w:style>',
    '<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:pPr><w:outlineLvl w:val="2"/></w:pPr><w:rPr><w:i/></w:rPr></w:style>',
  ].join("");
  const source = async (): Promise<Uint8Array> =>
    await buildSourceWith({
      body: [
        p(
          r("Comment: this clause is optional."),
          '<w:pStyle w:val="NoteText"/>',
        ),
        p(
          `${r("See ")}${r("Smith v. Jones", '<w:i w:val="0"/>')}${r(" for the rule.")}`,
          '<w:pStyle w:val="NoteText"/>',
        ),
        p(r("Optional Provisions"), '<w:pStyle w:val="Heading3"/>'),
      ].join(""),
      styles: STYLES,
    });
  const runsOf = (xml: string, text: string): string =>
    new RegExp(`<w:r>(?:(?!</w:r>).)*${text}`, "u").exec(xml)?.[0] ?? "";

  test("a note set in an italic text style stays italic on the template, a heading's italic is the template's", async () => {
    const applied = unwrap(
      await applyFor({
        source: await source(),
        template,
        items: styled("HouseBody"),
      }),
    );
    expect(applied.report.ok).toBe(true);
    const xml = await documentOf(applied.docx);
    expect(runsOf(xml, "Comment: this clause")).toContain("<w:i/>");
    expect(runsOf(xml, "See ")).toContain("<w:i/>");
    // A run that sets the style's italic off reads upright.
    expect(runsOf(xml, "Smith v. Jones")).not.toContain("<w:i/>");
    expect(runsOf(xml, "Optional Provisions")).not.toContain("<w:i/>");
  });
});

describe("page starts", () => {
  const PAGE_BREAK = '<w:r><w:br w:type="page"/></w:r>';
  const LANDSCAPE =
    '<w:sectPr><w:type w:val="continuous"/><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:top="1000" w:right="1100" w:bottom="1200" w:left="1300"/><w:cols w:num="2"/></w:sectPr>';

  test("the manifest says where the source starts a page or ends a section", async () => {
    const source = await buildSourceWith({
      body: [
        p(r("First.")),
        p(PAGE_BREAK),
        p(r("Annex"), "<w:pageBreakBefore/>"),
        p(r("By its style."), style("NewPage")),
        p(
          r("Switched off."),
          `${style("NewPage")}<w:pageBreakBefore w:val="0"/>`,
        ),
        p(
          `${r("Mid")}<w:r><w:br w:type="column"/></w:r>${PAGE_BREAK}`,
          LANDSCAPE,
        ),
        p(r("Last.")),
      ].join(""),
      styles:
        '<w:style w:type="paragraph" w:styleId="NewPage"><w:name w:val="New Page"/><w:pPr><w:pageBreakBefore/></w:pPr></w:style>',
    });
    const { blocks, sections } = unwrap(await analyzeDocxBlocks(source));

    const facts: unknown = blocks.map(
      ({ pageBreaks, pageBreakBefore, endsSection, sectionIndex }) => ({
        pageBreaks,
        pageBreakBefore,
        endsSection,
        sectionIndex,
      }),
    );
    expect(facts).toEqual([
      { sectionIndex: 0 },
      { pageBreaks: 1, sectionIndex: 0 },
      { pageBreakBefore: true, sectionIndex: 0 },
      { pageBreakBefore: true, sectionIndex: 0 },
      { sectionIndex: 0 },
      { pageBreaks: 1, endsSection: true, sectionIndex: 0 },
      { sectionIndex: 1 },
    ]);
    expect(sections).toEqual([
      {
        index: 0,
        start: "continuous",
        orientation: "landscape",
        size: { width: 16_838, height: 11_906 },
        margins: { top: 1000, right: 1100, bottom: 1200, left: 1300 },
        columns: 2,
      },
      {
        index: 1,
        start: "nextPage",
        orientation: "portrait",
        size: { width: null, height: null },
        margins: { top: null, right: null, bottom: null, left: null },
        columns: 1,
      },
    ]);
  });

  test("a body without section properties still ends in a section", async () => {
    const built = await buildMinimalDocx(p(r("Only.")));
    const bare = await replacePart(
      built,
      "word/document.xml",
      ((await readPart(built, "word/document.xml")) ?? "").replace(
        "<w:sectPr/>",
        "",
      ),
    );
    const { blocks, sections } = unwrap(await analyzeDocxBlocks(bare));

    expect(blocks.map(({ sectionIndex }) => sectionIndex)).toEqual([0]);
    expect(sections).toEqual([
      {
        index: 0,
        start: "nextPage",
        orientation: "portrait",
        size: { width: null, height: null },
        margins: { top: null, right: null, bottom: null, left: null },
        columns: 1,
      },
    ]);
  });

  const twoParagraphs = async (): Promise<Uint8Array> =>
    await buildMinimalDocx(`${p(r("Body."))}${p(r("Annex"))}`);

  test("an item starts its paragraph on a new page, between style and list", async () => {
    const { docx, report } = unwrap(
      await applyFor({
        source: await twoParagraphs(),
        template,
        items: (blocks) =>
          blocks.map(({ id, hash }, index) => ({
            kind: "source",
            blockId: id,
            hash,
            styleId: "HouseBody",
            ...(index === 1
              ? { pageBreakBefore: true, numbering: { numId: 1, ilvl: 0 } }
              : {}),
          })),
      }),
    );
    const document = await documentOf(docx);

    expect(report.ok).toBe(true);
    expect(report.counts.pageStarts).toBe(1);
    expect(document.split("<w:pageBreakBefore/>")).toHaveLength(2);
    expect(document).toContain(
      '<w:pPr><w:pStyle w:val="HouseBody"/><w:pageBreakBefore/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>',
    );
    expect(
      unwrap(await analyzeDocxBlocks(docx)).blocks.filter(
        ({ pageBreakBefore }) => pageBreakBefore === true,
      ),
    ).toHaveLength(1);
  });

  test("keeping with the next paragraph is verified by itself", async () => {
    const source = await twoParagraphs();
    const items = (keepNext: boolean) => (blocks: PlannedBlock[]) =>
      blocks.map(({ id, hash }, index): TemplatePlanItem => ({
        kind: "source",
        blockId: id,
        hash,
        styleId: "HouseBody",
        indent: { left: 851 },
        ...(index === 0 && keepNext ? { keepNext } : {}),
      }));
    const kept = await planFor({ source, template, items: items(true) });
    const loose = await planFor({ source, template, items: items(false) });
    // The same indent on both sides: only the keep-with-next differs.
    for (const [plan, other] of [
      [kept, loose],
      [loose, kept],
    ] as const) {
      const { docx } = unwrap(
        await applyTemplatePlan({ source, template, plan: other }),
      );
      expect(
        (await verifyFor({ source, template, output: docx, plan })).issues.map(
          ({ code, blockId }) => [code, blockId],
        ),
      ).toEqual([["paragraph-format-mismatch", "p0"]]);
    }
  });

  test("an item gives its paragraph tab stops of its own", async () => {
    const source = await twoParagraphs();
    const items =
      (tabStops: number[] | undefined) => (blocks: PlannedBlock[]) =>
        blocks.map(({ id, hash }, index): TemplatePlanItem => ({
          kind: "source",
          blockId: id,
          hash,
          styleId: "HouseBody",
          numbering: { numId: 1, ilvl: 1 },
          ...(index === 0 && tabStops !== undefined
            ? { tabStops, indent: { left: 851 } }
            : {}),
        }));
    const stopped = await planFor({
      source,
      template,
      items: items([2552, 2572]),
    });
    const plain = await planFor({ source, template, items: items(undefined) });
    const { docx, report } = unwrap(
      await applyTemplatePlan({ source, template, plan: stopped }),
    );
    const document = await documentOf(docx);

    expect(report.ok).toBe(true);
    // Behind the list and in front of the indent, as the schema orders them.
    expect(document).toContain(
      '<w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr><w:tabs><w:tab w:val="left" w:pos="2552"/><w:tab w:val="left" w:pos="2572"/></w:tabs><w:ind w:left="851"/></w:pPr>',
    );
    expect(document.split("<w:tabs>")).toHaveLength(2);
    // Verification holds the output to the plan both ways, and to each stop.
    const moved = await replacePart(
      docx,
      "word/document.xml",
      document.replace('w:pos="2572"', 'w:pos="3000"'),
    );
    for (const [plan, output] of [
      [plain, docx],
      [
        stopped,
        unwrap(await applyTemplatePlan({ source, template, plan: plain })).docx,
      ],
      [stopped, moved],
    ] as const) {
      expect(
        (await verifyFor({ source, template, output, plan })).issues.map(
          ({ code, blockId }) => [code, blockId],
        ),
      ).toEqual([["paragraph-format-mismatch", "p0"]]);
    }
  });

  test("an item gives its paragraph an alignment of its own, and a cell paragraph too", async () => {
    const source = await buildMinimalDocx(
      `${p(r("centred line"))}${p(r("plain line"))}<w:tbl><w:tr><w:tc>${p(r("amount"))}</w:tc></w:tr></w:tbl>${p(r("end"))}`,
    );
    const itemOf = (
      { id, hash }: PlannedBlock,
      aligned: boolean,
    ): TemplatePlanItem => {
      const item: TemplatePlanItem = { kind: "source", blockId: id, hash };
      if (id.startsWith("t")) {
        if (aligned) {
          item.nested = [{ blockId: "p3", alignment: "right" }];
        }
        return item;
      }
      item.styleId = "HouseBody";
      if (id === "p0" && aligned) {
        item.alignment = "center";
        item.indent = { left: 0 };
      }
      return item;
    };
    const items = (aligned: boolean) => (blocks: PlannedBlock[]) =>
      blocks
        .filter(({ id }) => id !== "p3")
        .map((block) => itemOf(block, aligned));
    const aligned = await planFor({ source, template, items: items(true) });
    const plain = await planFor({ source, template, items: items(false) });
    const { docx, report } = unwrap(
      await applyTemplatePlan({ source, template, plan: aligned }),
    );
    const document = await documentOf(docx);

    expect(report.ok).toBe(true);
    // Behind the indent, as the schema orders them.
    expect(document).toContain(
      '<w:pPr><w:pStyle w:val="HouseBody"/><w:ind w:left="0"/><w:jc w:val="center"/></w:pPr>',
    );
    expect(document).toContain(
      '<w:tc><w:p><w:pPr><w:jc w:val="right"/></w:pPr>',
    );
    expect(document.split("<w:jc ")).toHaveLength(3);
    // Verification holds the output to the plan both ways.
    const plainOutput = unwrap(
      await applyTemplatePlan({ source, template, plan: plain }),
    ).docx;
    expect(
      (
        await verifyFor({ source, template, output: docx, plan: plain })
      ).issues.map(({ code, blockId }) => [code, blockId]),
    ).toEqual([
      ["paragraph-format-mismatch", "p0"],
      ["paragraph-format-mismatch", "p3"],
    ]);
    expect(
      (
        await verifyFor({
          source,
          template,
          output: plainOutput,
          plan: aligned,
        })
      ).issues.map(({ code, blockId }) => [code, blockId]),
    ).toEqual([
      ["paragraph-format-mismatch", "p0"],
      ["paragraph-format-mismatch", "p3"],
    ]);
  });

  test("an item gives its paragraph a spacing of its own, and a cell paragraph too", async () => {
    const tight = '<w:spacing w:before="0" w:after="0"/>';
    const source = await buildMinimalDocx(
      [
        p(r("spaced line"), '<w:spacing w:before="240" w:after="60"/>'),
        p(r("plain line")),
        `<w:tbl><w:tr><w:tc>${p(r("swatch"), tight)}</w:tc></w:tr></w:tbl>`,
        // A side that follows the application, and one counted in lines,
        // state no measure.
        p(
          r("end"),
          '<w:spacing w:before="100" w:beforeAutospacing="1" w:after="120" w:afterLines="50"/>',
        ),
      ].join(""),
    );
    // The manifest says what each paragraph states for itself.
    const read = unwrap(await analyzeDocxBlocks(source)).blocks;
    expect(read.at(0)?.spacing).toEqual({ before: 240, after: 60 });
    expect(read.at(1)?.spacing).toBeUndefined();
    expect(read.at(2)?.nested?.at(0)?.spacing).toEqual({ before: 0, after: 0 });
    expect(read.at(3)?.spacing).toBeUndefined();

    const itemOf = (
      { id, hash }: PlannedBlock,
      spaced: boolean,
    ): TemplatePlanItem => {
      const item: TemplatePlanItem = { kind: "source", blockId: id, hash };
      if (id.startsWith("t")) {
        if (spaced) {
          item.nested = [{ blockId: "p3", spacing: { before: 0, after: 0 } }];
        }
        return item;
      }
      item.styleId = "HouseBody";
      if (id === "p0" && spaced) {
        item.spacing = { after: 60 };
        item.tabStops = [1000];
        item.indent = { left: 0 };
      }
      return item;
    };
    const items = (spaced: boolean) => (blocks: PlannedBlock[]) =>
      blocks
        .filter(({ id }) => id !== "p3")
        .map((block) => itemOf(block, spaced));
    const spaced = await planFor({ source, template, items: items(true) });
    const plain = await planFor({ source, template, items: items(false) });
    const { docx, report } = unwrap(
      await applyTemplatePlan({ source, template, plan: spaced }),
    );
    const document = await documentOf(docx);

    expect(report.ok).toBe(true);
    // Between the tab stops and the indent, as the schema orders them; a
    // side the plan leaves out is the style's.
    expect(document).toContain(
      '<w:pPr><w:pStyle w:val="HouseBody"/><w:tabs><w:tab w:val="left" w:pos="1000"/></w:tabs><w:spacing w:after="60"/><w:ind w:left="0"/></w:pPr>',
    );
    expect(document).toContain(
      '<w:tc><w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr>',
    );
    expect(document.split("<w:spacing ")).toHaveLength(3);
    // Verification holds the output to the plan both ways.
    const plainOutput = unwrap(
      await applyTemplatePlan({ source, template, plan: plain }),
    ).docx;
    for (const [output, plan] of [
      [docx, plain],
      [plainOutput, spaced],
    ] as const) {
      expect(
        (await verifyFor({ source, template, output, plan })).issues.map(
          ({ code, blockId }) => [code, blockId],
        ),
      ).toEqual([
        ["paragraph-format-mismatch", "p0"],
        ["paragraph-format-mismatch", "p3"],
      ]);
    }
    // A line spacing is written with its rule, and held to the plan.
    const lined = await planFor({
      source,
      template,
      items: (blocks) =>
        blocks
          .filter(({ id }) => id !== "p3")
          .map(({ id, hash }): TemplatePlanItem => {
            const item: TemplatePlanItem = {
              kind: "source",
              blockId: id,
              hash,
            };
            if (id === "p0") {
              item.spacing = { after: 0, line: 276 };
            }
            if (id === "p1") {
              item.spacing = { line: 280, lineRule: "exact" };
            }
            return item;
          }),
    });
    const linedOutput = unwrap(
      await applyTemplatePlan({ source, template, plan: lined }),
    ).docx;
    const linedDocument =
      (await readPart(linedOutput, "word/document.xml")) ?? "";
    expect(linedDocument).toContain(
      '<w:spacing w:after="0" w:line="276" w:lineRule="auto"/>',
    );
    expect(linedDocument).toContain(
      '<w:spacing w:line="280" w:lineRule="exact"/>',
    );
    expect(
      (await verifyFor({ source, template, output: linedOutput, plan: lined }))
        .issues,
    ).toEqual([]);
    // A spacing without a side, with another key, with a fraction or below
    // zero, a line spacing of none, a rule without a line spacing or a rule
    // the format does not know is no spacing a plan may state.
    for (const spacing of [
      {},
      { gap: 240 },
      { before: 1.5 },
      { after: -20 },
      { line: 0 },
      { lineRule: "exact" },
      { line: 240, lineRule: "double" },
    ]) {
      const malformed = await planFor({
        source,
        template,
        items: (blocks) =>
          blocks
            .filter(({ id }) => id !== "p3")
            .map(({ id, hash }): TemplatePlanItem => {
              const item: TemplatePlanItem = {
                kind: "source",
                blockId: id,
                hash,
              };
              if (id === "p0") {
                Object.assign(item, { spacing });
              }
              return item;
            }),
      });
      expect(
        codeOf(await applyTemplatePlan({ source, template, plan: malformed })),
      ).toBe("plan-invalid");
    }
  });

  test("an item keeps its paragraph with the next and gives it an indent of its own", async () => {
    const source = await twoParagraphs();
    const items = (formatted: boolean) => (blocks: PlannedBlock[]) =>
      blocks.map(({ id, hash }, index): TemplatePlanItem => ({
        kind: "source",
        blockId: id,
        hash,
        styleId: "HouseBody",
        ...(index === 0 && formatted
          ? { keepNext: true, indent: { left: 851, hanging: 851 } }
          : {}),
      }));
    const formatted = await planFor({ source, template, items: items(true) });
    const plain = await planFor({ source, template, items: items(false) });
    const { docx, report } = unwrap(
      await applyTemplatePlan({ source, template, plan: formatted }),
    );
    const document = await documentOf(docx);

    expect(report.ok).toBe(true);
    // In the order the schema gives a paragraph's properties.
    expect(document).toContain(
      '<w:pPr><w:pStyle w:val="HouseBody"/><w:keepNext/><w:ind w:left="851" w:hanging="851"/></w:pPr>',
    );
    expect(document.split("<w:keepNext/>")).toHaveLength(2);
    // Verification holds the output to the plan both ways.
    expect(
      (
        await verifyFor({ source, template, output: docx, plan: plain })
      ).issues.map(({ code, blockId }) => [code, blockId]),
    ).toEqual([["paragraph-format-mismatch", "p0"]]);
    const unformatted = unwrap(
      await applyTemplatePlan({ source, template, plan: plain }),
    );
    expect(
      (
        await verifyFor({
          source,
          template,
          output: unformatted.docx,
          plan: formatted,
        })
      ).issues.map(({ code, blockId }) => [code, blockId]),
    ).toEqual([["paragraph-format-mismatch", "p0"]]);
    // The same indent written with a unit is the same indent; one counted
    // in characters beside it is another paragraph.
    const written = '<w:ind w:left="851" w:hanging="851"/>';
    const outputWith = async (indent: string): Promise<Uint8Array> =>
      await replacePart(
        docx,
        "word/document.xml",
        document.replace(written, indent),
      );
    expect(
      (
        await verifyFor({
          source,
          template,
          plan: formatted,
          output: await outputWith(
            '<w:ind w:left="42.55pt" w:hanging="42.55pt"/>',
          ),
        })
      ).issues.map(({ code }) => code),
    ).toEqual([]);
    for (const changed of [
      '<w:ind w:left="1in" w:hanging="851"/>',
      '<w:ind w:left="851" w:leftChars="200" w:hanging="851"/>',
    ]) {
      expect(
        (
          await verifyFor({
            source,
            template,
            plan: formatted,
            output: await outputWith(changed),
          })
        ).issues.map(({ code, blockId }) => [code, blockId]),
      ).toEqual([["paragraph-format-mismatch", "p0"]]);
    }
    // A hanging and a first-line indent exclude each other.
    expect(
      codeOf(
        await applyFor({
          source,
          template,
          items: (blocks) =>
            blocks.map(({ id, hash }) => ({
              kind: "source",
              blockId: id,
              hash,
              styleId: "HouseBody",
              indent: { hanging: 300, firstLine: 300 },
            })),
        }),
      ),
    ).toBe("plan-invalid");
  });

  test("verification notices a page start the plan does not give, and a missing one", async () => {
    const source = await twoParagraphs();
    const items = (pageBreakBefore: boolean) => (blocks: PlannedBlock[]) =>
      blocks.map(({ id, hash }, index): TemplatePlanItem => ({
        kind: "source",
        blockId: id,
        hash,
        styleId: "HouseBody",
        ...(index === 1 && pageBreakBefore ? { pageBreakBefore } : {}),
      }));
    const withStart = await planFor({ source, template, items: items(true) });
    const without = await planFor({ source, template, items: items(false) });
    const { docx } = unwrap(
      await applyTemplatePlan({ source, template, plan: withStart }),
    );

    expect(
      (
        await verifyFor({ source, template, output: docx, plan: without })
      ).issues.map(({ code, blockId }) => [code, blockId]),
    ).toEqual([["page-start-mismatch", "p1"]]);
    // One that is there and switched off starts no page.
    const switchedOff = await replacePart(
      docx,
      "word/document.xml",
      (await documentOf(docx)).replace(
        "<w:pageBreakBefore/>",
        '<w:pageBreakBefore w:val="0"/>',
      ),
    );
    expect(
      (
        await verifyFor({
          source,
          template,
          output: switchedOff,
          plan: withStart,
        })
      ).issues.map(({ code, blockId }) => [code, blockId]),
    ).toEqual([["page-start-mismatch", "p1"]]);
    const plain = unwrap(
      await applyTemplatePlan({ source, template, plan: without }),
    );
    expect(
      (
        await verifyFor({
          source,
          template,
          output: plain.docx,
          plan: withStart,
        })
      ).issues.map(({ code, blockId }) => [code, blockId]),
    ).toEqual([["page-start-mismatch", "p1"]]);
  });

  test("a table, a joined block and a malformed value cannot take a page start", async () => {
    const table = await buildMinimalDocx(
      `<w:tbl><w:tr><w:tc>${p(r("cell"))}</w:tc></w:tr></w:tbl>${p(r("after"))}`,
    );
    expect(
      codeOf(
        await applyFor({
          source: table,
          template,
          items: (blocks) =>
            blocks.map(({ id, hash }, index) => ({
              kind: "source",
              blockId: id,
              hash,
              ...(index === 0 ? { pageBreakBefore: true } : {}),
            })),
        }),
      ),
    ).toBe("plan-invalid");
    expect(
      codeOf(
        await applyFor({
          source: await twoParagraphs(),
          template,
          items: (blocks) =>
            blocks.map(({ id, hash }, index) => ({
              kind: "source",
              blockId: id,
              hash,
              ...(index === 0
                ? { mergeNext: "space" }
                : { pageBreakBefore: true }),
            })),
        }),
      ),
    ).toBe("merge-invalid");
    const malformed: TemplatePlan = JSON.parse(
      JSON.stringify(
        await planFor({ source: await twoParagraphs(), template }),
      ).replaceAll(
        '"kind":"source"',
        '"kind":"source","pageBreakBefore":"yes"',
      ),
    );
    expect(
      codeOf(
        await applyTemplatePlan({
          source: await twoParagraphs(),
          template,
          plan: malformed,
        }),
      ),
    ).toBe("plan-invalid");
  });
});

describe("justified lines in front of a manual line break", () => {
  const settings = (inner: string): string =>
    `${XML}<w:settings ${NS}>${inner}</w:settings>`;

  test("the setting is written at its place, and nothing else changes", () => {
    for (const [before, after] of [
      // No compatibility settings yet: they stand in front of the revision ids.
      [
        '<w:zoom w:percent="100"/><w:rsids><w:rsidRoot w:val="00000001"/></w:rsids>',
        '<w:zoom w:percent="100"/><w:compat><w:doNotExpandShiftReturn/></w:compat><w:rsids><w:rsidRoot w:val="00000001"/></w:rsids>',
      ],
      // Among the other settings in the schema's order.
      [
        '<w:compat><w:spaceForUL/><w:ulTrailSpace/><w:useFELayout/><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>',
        '<w:compat><w:spaceForUL/><w:ulTrailSpace/><w:doNotExpandShiftReturn/><w:useFELayout/><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>',
      ],
      ["<w:compat/>", "<w:compat><w:doNotExpandShiftReturn/></w:compat>"],
      [
        '<w:compat><w:doNotExpandShiftReturn w:val="0"/><w:useFELayout/></w:compat>',
        "<w:compat><w:doNotExpandShiftReturn/><w:useFELayout/></w:compat>",
      ],
      // Already on: not a byte changes.
      [
        '<w:compat><w:doNotExpandShiftReturn w:val="true"/></w:compat>',
        '<w:compat><w:doNotExpandShiftReturn w:val="true"/></w:compat>',
      ],
    ] as const) {
      const written = setKeepShiftReturnSpacing(settings(before));
      expect(written).toBe(settings(after));
      expect(keepsShiftReturnSpacing(written ?? "")).toBe(true);
    }
    expect(keepsShiftReturnSpacing(settings("<w:compat/>"))).toBe(false);
    expect(setKeepShiftReturnSpacing(`${XML}<w:settings ${NS}/>`)).toBe(
      settings("<w:compat><w:doNotExpandShiftReturn/></w:compat>"),
    );
    expect(setKeepShiftReturnSpacing(`${XML}<w:styles ${NS}/>`)).toBeNull();
  });

  test("a plan switches the setting on in the output, and only when it asks", async () => {
    const source = await buildMinimalDocx(p(r("Body.")));
    const plain = unwrap(await applyFor({ source, template }));
    const asked = unwrap(
      await applyFor({
        source,
        template,
        extra: { compat: { doNotExpandShiftReturn: true } },
      }),
    );
    const before = (await readPart(plain.docx, "word/settings.xml")) ?? "";
    const after = (await readPart(asked.docx, "word/settings.xml")) ?? "";

    expect(keepsShiftReturnSpacing(before)).toBe(false);
    expect(keepsShiftReturnSpacing(after)).toBe(true);
    expect(after.replace("<w:doNotExpandShiftReturn/>", "")).toBe(before);
    expect(unwrap(await analyzeDocxBlocks(asked.docx)).compat).toEqual({
      doNotExpandShiftReturn: true,
    });

    // The plan asks, the output does not have it.
    const plan = await planFor({
      source,
      template,
      extra: { compat: { doNotExpandShiftReturn: true } },
    });
    expect(
      (
        await verifyFor({ source, template, output: plain.docx, plan })
      ).issues.map(({ code }) => code),
    ).toEqual(["compat-mismatch"]);
  });

  test("the manifest says how each paragraph is aligned", async () => {
    const source = await buildSourceWith({
      body: [
        p(r("By its style.")),
        p(r("By its own."), '<w:jc w:val="center"/>'),
        p(r("By the style it builds on."), style("Derived")),
        p(r("Unstated."), style("Plain")),
      ].join(""),
      styles:
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr><w:jc w:val="both"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="Derived"><w:name w:val="Derived"/><w:basedOn w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Plain"><w:name w:val="Plain"/></w:style>',
    });
    const { blocks, compat } = unwrap(await analyzeDocxBlocks(source));

    expect(blocks.map(({ alignment }) => alignment ?? null)).toEqual([
      "both",
      "center",
      "both",
      null,
    ]);
    expect(compat).toEqual({ doNotExpandShiftReturn: false });
  });
});

describe("tables kept apart", () => {
  const table = (text: string): string =>
    `<w:tbl><w:tr><w:tc>${p(r(text))}</w:tc></w:tr></w:tbl>`;

  test("an empty paragraph parts two tables a plan puts next to each other and closes the body after one", async () => {
    const source = await buildMinimalDocx(
      `${p(r("Intro."))}${table("one")}${p("")}${table("two")}${p("")}`,
    );
    const { docx, report } = unwrap(
      await applyFor({
        source,
        template,
        items: (blocks) =>
          blocks
            .filter(({ id }) => id !== "p3" && id !== "p6")
            .map(({ id, hash }) => ({ kind: "source", blockId: id, hash })),
        extra: {
          omit: [
            { blockId: "p3", reason: "empty" },
            { blockId: "p6", reason: "empty" },
          ],
        },
      }),
    );
    const document = await documentOf(docx);

    expect(report.ok).toBe(true);
    expect(report.counts.tableSeparators).toBe(2);
    expect(document).not.toContain("</w:tbl><w:tbl>");
    expect(document.split("</w:tbl><w:p/>")).toHaveLength(3);
    expect(document).toContain("</w:tbl><w:p/><w:sectPr");
  });

  test("tables that a paragraph parts get none, and verification misses one that is gone", async () => {
    const source = await buildMinimalDocx(
      `${table("one")}${p(r("Between."))}${table("two")}`,
    );
    const plan = await planFor({ source, template });
    const { docx, report } = unwrap(
      await applyTemplatePlan({ source, template, plan }),
    );
    const document = await documentOf(docx);

    expect(report.counts.tableSeparators).toBe(1);
    const closed = "</w:tbl><w:p/><w:sectPr";
    expect(document).toContain(closed);
    const open = await verifyFor({
      source,
      template,
      output: await replacePart(
        docx,
        "word/document.xml",
        document.replace(closed, "</w:tbl><w:sectPr"),
      ),
      plan,
    });
    expect(open.issues.map(({ code }) => code)).toContain("block-missing");
  });
});

describe("sections a plan ends", () => {
  const LANDSCAPE = {
    orientation: "landscape",
    width: 16_838,
    height: 11_906,
    margins: { top: 1000, right: 1100, bottom: 1200, left: 1300 },
  } as const;
  const three = async (): Promise<Uint8Array> =>
    await buildMinimalDocx(
      `${p(r("Body."))}${p(r("Wide annex."))}${p(r("After."))}`,
    );
  const ending =
    (sections: Record<number, TemplatePlanSourceItem["endsSection"]>) =>
    (blocks: PlannedBlock[]): TemplatePlanItem[] =>
      blocks.map(({ id, hash }, index) => ({
        kind: "source",
        blockId: id,
        hash,
        ...(sections[index] === undefined
          ? {}
          : { endsSection: sections[index] }),
      }));
  const pages = async (docx: Uint8Array) =>
    unwrap(await analyzeDocxBlocks(docx)).sections.map(
      ({ start, orientation, size, margins }) => [
        start,
        orientation,
        size.width,
        margins.top,
      ],
    );

  test("a section is the template's last one with the plan's page in place", async () => {
    const { docx, report } = unwrap(
      await applyFor({
        source: await three(),
        template,
        items: ending({ 0: {}, 1: { page: LANDSCAPE } }),
      }),
    );
    const document = await documentOf(docx);
    const sections = [...document.matchAll(/<w:sectPr>.*?<\/w:sectPr>/gu)].map(
      ([xml]) => xml,
    );

    expect(report.ok).toBe(true);
    expect(report.counts.sectionEnds).toBe(2);
    expect(await pages(docx)).toEqual([
      ["nextPage", "portrait", 11_906, 1417],
      ["nextPage", "landscape", 16_838, 1000],
      ["nextPage", "portrait", 11_906, 1417],
    ]);
    // Each carries the template's header and footer.
    expect(sections).toHaveLength(3);
    for (const section of sections) {
      expect(section).toContain("<w:headerReference ");
      expect(section).toContain("<w:footerReference ");
    }
    expect(sections.at(1)).toContain(
      '<w:type w:val="nextPage"/><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:top="1000" w:right="1100" w:bottom="1200" w:left="1300"',
    );
    // The section stands at the end of its paragraph's properties.
    expect(document).toContain("</w:sectPr></w:pPr><w:r>");
    expect(
      unwrap(await analyzeDocxBlocks(docx))
        .blocks.filter(({ endsSection }) => endsSection === true)
        .map(({ text }) => text),
    ).toEqual(["Body.", "Wide annex."]);
  });

  test("the body can end in a section of another page", async () => {
    const { docx, report } = unwrap(
      await applyFor({
        source: await three(),
        template,
        items: ending({ 1: {} }),
        extra: { lastSection: { page: LANDSCAPE, start: "oddPage" } },
      }),
    );

    expect(report.ok).toBe(true);
    expect(await pages(docx)).toEqual([
      ["nextPage", "portrait", 11_906, 1417],
      ["oddPage", "landscape", 16_838, 1000],
    ]);
  });

  test("only the document's first section keeps the first-page header and the page-number start", () => {
    const part = readRawPart(
      "word/document.xml",
      `<w:document ${NS}><w:body><w:sectPr w:rsidR="00AA00AA"><w:headerReference w:type="default" r:id="rId7"/><w:type w:val="continuous"/><w:pgSz w:w="11906" w:h="16838" w:code="9"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1134" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/><w:pgNumType w:fmt="lowerRoman" w:start="1"/><w:cols w:space="708"/><w:titlePg/><w:docGrid w:linePitch="360"/></w:sectPr></w:body></w:document>`,
    );
    const sectPr = part?.root.elements
      ?.at(0)
      ?.elements?.find(({ name }) => name === "w:sectPr");
    if (part === null || part === undefined || sectPr === undefined) {
      throw new Error("fixture");
    }
    const rest =
      '<w:cols w:space="708"/><w:titlePg/><w:docGrid w:linePitch="360"/>';

    expect(
      sectionPropertiesXml(part.xml, sectPr, "w", {
        first: true,
        page: undefined,
        start: null,
      }),
    ).toBe(
      `<w:sectPr><w:headerReference w:type="default" r:id="rId7"/><w:type w:val="continuous"/><w:pgSz w:w="11906" w:h="16838" w:code="9"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1134" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/><w:pgNumType w:fmt="lowerRoman" w:start="1"/>${rest}</w:sectPr>`,
    );
    expect(
      sectionPropertiesXml(part.xml, sectPr, "w", {
        first: false,
        page: LANDSCAPE,
        start: "nextPage",
      }),
    ).toBe(
      '<w:sectPr><w:headerReference w:type="default" r:id="rId7"/><w:type w:val="nextPage"/><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:top="1000" w:right="1100" w:bottom="1200" w:left="1300" w:header="708" w:footer="708" w:gutter="0"/><w:pgNumType w:fmt="lowerRoman"/><w:cols w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr>',
    );
  });

  test("a section built from the template's last one states the headers that one inherits", async () => {
    // The template's first section states the header and the footer; its
    // last section states none and shows the same.
    const base = await buildTemplate();
    const document = (await readPart(base, "word/document.xml")) ?? "";
    const stated =
      /<w:sectPr>(?<references>(?:<w:(?:header|footer)Reference [^>]*\/>)+)(?<rest>.*?)<\/w:sectPr><\/w:body>/u.exec(
        document,
      )?.groups;
    if (stated === undefined) {
      throw new Error("fixture");
    }
    const references = stated["references"] ?? "";
    const rest = stated["rest"] ?? "";
    const inheriting = await replacePart(
      base,
      "word/document.xml",
      document
        .replace(
          "<w:body>",
          `<w:body><w:p><w:pPr><w:sectPr>${references}${rest}</w:sectPr></w:pPr></w:p>`,
        )
        .replace(
          `<w:sectPr>${references}${rest}</w:sectPr></w:body>`,
          `<w:sectPr>${rest}</w:sectPr></w:body>`,
        ),
    );
    expect(inheriting).not.toBe(base);
    const source = await three();
    const whole = await planFor({
      source,
      template: inheriting,
      items: ending({}),
    });
    const parted = await planFor({
      source,
      template: inheriting,
      items: ending({ 0: {} }),
    });
    for (const plan of [whole, parted]) {
      const { docx, report } = unwrap(
        await applyTemplatePlan({ source, template: inheriting, plan }),
      );
      expect(report.ok).toBe(true);
      const sections = [
        ...((await readPart(docx, "word/document.xml")) ?? "").matchAll(
          /<w:sectPr>.*?<\/w:sectPr>/gu,
        ),
      ].map(([section]) => section);
      expect(sections.length).toBeGreaterThan(0);
      for (const section of sections) {
        expect(section.startsWith(`<w:sectPr>${references}`)).toBe(true);
      }
    }
  });

  test("verification holds the output to the plan's sections", async () => {
    const source = await three();
    const plan = await planFor({
      source,
      template,
      items: ending({ 0: {}, 1: { page: LANDSCAPE } }),
    });
    const { docx } = unwrap(
      await applyTemplatePlan({ source, template, plan }),
    );
    const document = await documentOf(docx);
    const portrait = await replacePart(
      docx,
      "word/document.xml",
      document.replace(
        '<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>',
        '<w:pgSz w:w="11906" w:h="16838"/>',
      ),
    );
    expect(
      (
        await verifyFor({ source, template, output: portrait, plan })
      ).issues.map(({ code }) => code),
    ).toEqual(["section-mismatch"]);

    const plain = await planFor({ source, template });
    expect(
      (
        await verifyFor({ source, template, output: docx, plan: plain })
      ).issues.map(({ code, blockId }) => [code, blockId]),
    ).toEqual([
      ["section-mismatch", "p0"],
      ["section-mismatch", "p1"],
    ]);
  });

  test("a table, a joined block and a malformed page cannot end a section", async () => {
    const table = await buildMinimalDocx(
      `<w:tbl><w:tr><w:tc>${p(r("cell"))}</w:tc></w:tr></w:tbl>${p(r("after"))}`,
    );
    expect(
      codeOf(
        await applyFor({ source: table, template, items: ending({ 0: {} }) }),
      ),
    ).toBe("plan-invalid");
    expect(
      codeOf(
        await applyFor({
          source: await three(),
          template,
          items: (blocks) =>
            blocks.map(({ id, hash }, index) => ({
              kind: "source",
              blockId: id,
              hash,
              ...(index === 0 ? { mergeNext: "space" } : {}),
              ...(index === 1 ? { endsSection: {} } : {}),
            })),
        }),
      ),
    ).toBe("merge-invalid");
    const malformed: TemplatePlan = JSON.parse(
      JSON.stringify(
        await planFor({
          source: await three(),
          template,
          items: ending({ 0: { page: LANDSCAPE } }),
        }),
      ).replace('"width":16838', '"width":"wide"'),
    );
    expect(
      codeOf(
        await applyTemplatePlan({
          source: await three(),
          template,
          plan: malformed,
        }),
      ),
    ).toBe("plan-invalid");
  });
});

describe("omitted blocks", () => {
  const NUMBERING =
    '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>';
  const source = async (): Promise<Uint8Array> =>
    await buildSourceWith({
      body: [
        p(r("Text.")),
        p('<w:r><w:br w:type="page"/></w:r>'),
        p("", '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'),
        p(
          "",
          '<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/></w:sectPr>',
        ),
      ].join(""),
      numbering: NUMBERING,
    });
  const onlyText = (blocks: PlannedBlock[]): TemplatePlanItem[] =>
    blocks
      .slice(0, 1)
      .map(({ id, hash }) => ({ kind: "source", blockId: id, hash }));

  test("an empty paragraph that shows a number is not dropped without a word", async () => {
    const applied = await applyFor({
      source: await source(),
      template,
      items: onlyText,
    });

    expect(codeOf(applied)).toBe("block-unplaced");
    if (Result.isError(applied)) {
      expect(applied.error.blockId).toBe("p2");
    }
  });

  test("breaks and labels in omitted blocks are reported", async () => {
    const { report } = unwrap(
      await applyFor({
        source: await source(),
        template,
        items: onlyText,
        extra: { omit: [{ blockId: "p2", reason: "empty" }] },
      }),
    );
    const warnings = report.warnings.join(" ");

    expect(report.ok).toBe(true);
    expect(report.dropped.sectionBreaks).toBe(1);
    expect(warnings).toContain("1 page or column break(s)");
    expect(warnings).toContain("1 omitted block(s) showed a number");
    expect(warnings).toContain("1 section break(s)");
  });
});

describe("tables", () => {
  const cell = (inner: string): string => `<w:tc>${inner}</w:tc>`;

  test("a cell that ends in a table gets its closing paragraph and no revision record", async () => {
    const source = await buildMinimalDocx(
      `<w:tbl><w:tblGrid><w:gridCol w:w="3000"/><w:tblGridChange w:id="9"><w:tblGrid><w:gridCol w:w="1"/></w:tblGrid></w:tblGridChange></w:tblGrid><w:tr>${cell(`${p(r("outer"))}<w:tbl><w:tr>${cell(p(r("inner")))}</w:tr></w:tbl>`)}</w:tr></w:tbl>`,
    );
    const { docx, report } = unwrap(await applyFor({ source, template }));
    const document = await documentOf(docx);

    expect(report.ok).toBe(true);
    expect(document).not.toContain("tblGridChange");
    expect(document).toContain("</w:tbl><w:p/></w:tc>");
    expect(
      report.blocks.map(({ blockId, outputBlockId }) => [
        blockId,
        outputBlockId,
      ]),
    ).toEqual([
      ["t0", "t0"],
      ["p1", "p1"],
      ["t2", "t2"],
      ["p3", "p3"],
    ]);
  });

  test("a table that keeps its own borders keeps the fill of its cells, by the colour stated", async () => {
    const shaded = (properties: string): string =>
      `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/>${properties}</w:tcPr>${p(r("cell"))}</w:tc>`;
    const source = await buildMinimalDocx(
      `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tr>${shaded('<w:shd w:val="clear" w:color="auto" w:fill="E7E6E6" w:themeFill="background2"/><w:vAlign w:val="center"/>')}${shaded('<w:shd w:val="clear" w:color="auto" w:fill="auto"/>')}${shaded('<w:shd w:val="clear" w:themeFill="accent1"/>')}</w:tr></w:tbl>`,
    );
    const { docx, report } = unwrap(await applyFor({ source, template }));
    const document = await documentOf(docx);

    expect(report.ok).toBe(true);
    expect(document).toContain(
      '<w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="E7E6E6"/><w:vAlign w:val="center"/></w:tcPr>',
    );
    expect(document.match(/<w:shd /gu)).toHaveLength(1);
    expect(document).not.toContain("themeFill");
  });

  test("a plan empties a sample paragraph of a kept template table, and only such a one", async () => {
    const source = await buildMinimalDocx(p(r("Body text.")));
    const items = (blocks: PlannedBlock[]): TemplatePlanItem[] => [
      { kind: "template", templateBlockIds: [TEMPLATE_IDS.signatures] },
      ...blocks.map(({ id, hash }): TemplatePlanItem => ({
        kind: "source",
        blockId: id,
        hash,
        styleId: "HouseBody",
      })),
    ];
    const emptying = await planFor({
      source,
      template,
      items,
      extra: { emptyTemplateBlockIds: [TEMPLATE_IDS.signatureRight] },
    });
    const { docx, report } = unwrap(
      await applyTemplatePlan({ source, template, plan: emptying }),
    );
    const document = await documentOf(docx);

    expect(report.ok).toBe(true);
    // The left cell keeps its sample; the right one is a paragraph without runs.
    expect(document).toContain(
      '<w:tc><w:tcPr><w:tcW w:w="4531" w:type="dxa"/></w:tcPr><w:p></w:p></w:tc></w:tr></w:tbl>',
    );
    // Verification holds the output to it: the sample back in is a mismatch.
    const plain = unwrap(
      await applyTemplatePlan({
        source,
        template,
        plan: await planFor({ source, template, items }),
      }),
    ).docx;
    expect(
      (
        await verifyFor({ source, template, output: plain, plan: emptying })
      ).issues.map(({ code }) => code),
    ).toEqual(["template-mismatch"]);
    // Not a block outside a kept table, not one a slot fills, not twice.
    for (const ids of [
      [TEMPLATE_IDS.title],
      ["p999"],
      [TEMPLATE_IDS.signatureRight, TEMPLATE_IDS.signatureRight],
    ]) {
      expect(
        codeOf(
          await applyFor({
            source,
            template,
            items,
            extra: { emptyTemplateBlockIds: ids },
          }),
        ),
      ).toBe("unknown-block");
    }
  });

  test("verification notices cells folded into one", async () => {
    const source = await buildMinimalDocx(
      `<w:tbl><w:tr>${cell(p(r("left")))}${cell(p(r("right")))}</w:tr></w:tbl>`,
    );
    const plan = await planFor({ source, template });
    const { docx } = unwrap(
      await applyTemplatePlan({ source, template, plan }),
    );
    const document = await documentOf(docx);
    const folded = document.replace("</w:tc><w:tc>", "");
    expect(folded).not.toBe(document);
    const report = await verifyFor({
      source,
      template,
      output: await replacePart(docx, "word/document.xml", folded),
      plan,
    });

    expect(report.issues.map(({ code }) => code)).toEqual(["block-missing"]);
  });
});

describe("verification of a written output", () => {
  let source: Uint8Array;
  let plan: TemplatePlan;
  let docx: Uint8Array;

  beforeAll(async () => {
    source = await buildSourceWith({
      body: [
        p(`${r("See part 1")}${NOTE_REFERENCE}${r(".")}`),
        p(`${r("Target")}<w:r><w:br w:type="page"/></w:r>`),
      ].join(""),
      footnotes: note(p(r("The note."))),
    });
    plan = await planFor({
      source,
      template,
      extra: {
        references: [
          {
            blockId: "p0",
            start: 9,
            end: 10,
            targetBlockId: "p1",
            instruction: "\\h",
            expectedText: "1",
          },
        ],
      },
    });
    docx = unwrap(await applyTemplatePlan({ source, template, plan })).docx;
  });

  const tampered = async (edit: (xml: string) => string): Promise<string[]> => {
    const document = await documentOf(docx);
    const edited = edit(document);
    expect(edited).not.toBe(document);
    const report = await verifyFor({
      source,
      template,
      output: await replacePart(docx, "word/document.xml", edited),
      plan,
    });
    return report.issues.map(({ code }) => code);
  };

  test("the untouched output verifies", async () => {
    expect((await verifyFor({ source, template, output: docx, plan })).ok).toBe(
      true,
    );
  });

  test("a character style slipped onto a run counts as emphasis", async () => {
    expect(
      await tampered((xml) =>
        xml.replace(
          '<w:r><w:t xml:space="preserve">See part </w:t></w:r>',
          '<w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t xml:space="preserve">See part </w:t></w:r>',
        ),
      ),
    ).toEqual(["emphasis-mismatch"]);
  });

  test("a dropped page break is noticed", async () => {
    expect(
      await tampered((xml) =>
        xml.replace('<w:r><w:br w:type="page"/></w:r>', ""),
      ),
    ).toEqual(["text-mismatch"]);
  });

  test("a bookmark that ends before it starts is noticed", async () => {
    expect(
      await tampered((xml) =>
        xml.replace(
          /(<w:bookmarkStart w:id="\d+" w:name="_Ref\d{9}"\/>)(.*?)(<w:bookmarkEnd w:id="\d+"\/>)/u,
          "$3$2$1",
        ),
      ),
    ).toContain("bookmark-dangling");
  });

  test("a note reference that leads to a separator is noticed", async () => {
    expect(
      await tampered((xml) =>
        xml.replace(
          '<w:footnoteReference w:id="1"/>',
          '<w:footnoteReference w:id="0"/>',
        ),
      ),
    ).toEqual(expect.arrayContaining(["note-mismatch", "note-dangling"]));
  });
});

describe("the fixture source still applies whole", () => {
  test("every block placed with defaults reads back identical", async () => {
    const source = await buildSource();
    const { report } = unwrap(await applyFor({ source, template }));

    expect(report.ok).toBe(true);
    expect(report.counts.textMismatch).toBe(0);
  });
});

describe("checkbox content controls", () => {
  const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
  const box = (id: number, checked: 0 | 1): string =>
    `<w:sdt><w:sdtPr><w:id w:val="${String(id)}"/><w14:checkbox xmlns:w14="${W14}"><w14:checked w14:val="${String(checked)}"/><w14:checkedState w14:val="2612" w14:font="MS Gothic"/><w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/></w14:checkbox></w:sdtPr><w:sdtEndPr/><w:sdtContent>${r(checked === 1 ? "\u2612" : "\u2610")}</w:sdtContent></w:sdt>`;
  const BODY = [
    p(
      `${box(11, 0)}${r(" If this box is checked, the appendix is not used.")}`,
    ),
    p(`${r("Agreed: ")}${box(12, 1)}`),
    // A plain character that only looks like a box.
    p(r("\u2610 Not a control.")),
  ].join("");

  test("the mark of a checkbox goes back into its control, in the body and in a slot", async () => {
    const source = await buildMinimalDocx(BODY);
    const plan = await planFor({
      source,
      template,
      items: ([first, second, third]) => [
        { kind: "template", templateBlockIds: [TEMPLATE_IDS.signatures] },
        {
          kind: "slot",
          templateBlockId: TEMPLATE_IDS.signatureLeft,
          sourceBlockIds: second === undefined ? [] : [second.id],
        },
        ...[first, third].flatMap((block): TemplatePlanItem[] =>
          block === undefined
            ? []
            : [
                {
                  kind: "source",
                  blockId: block.id,
                  hash: block.hash,
                  styleId: "HouseBody",
                },
              ],
        ),
      ],
    });
    const { docx, report } = unwrap(
      await applyTemplatePlan({ source, template, plan }),
    );
    const document = await documentOf(docx);

    expect(report.ok).toBe(true);
    // Both controls with their properties, the text as it was.
    expect(document.split("<w:sdt>")).toHaveLength(3);
    expect(document).toMatch(
      /<w:sdt><w:sdtPr[^>]*><w:id w:val="11"\/><w14:checkbox[^>]*><w14:checked w14:val="0"\/>.*?<\/w:sdtPr><w:sdtContent><w:r>(?:<w:rPr>.*?<\/w:rPr>)?<w:t[^>]*>\u2610<\/w:t><\/w:r><\/w:sdtContent><\/w:sdt>/u,
    );
    expect(document).toMatch(
      /Agreed: <\/w:t><\/w:r><w:sdt><w:sdtPr[^>]*><w:id w:val="12"\/><w14:checkbox[^>]*><w14:checked w14:val="1"\/>/u,
    );
    const texts = unwrap(await analyzeDocxBlocks(docx)).blocks.flatMap(
      (block) => [block, ...(block.nested ?? [])].map(({ text }) => text),
    );
    expect(texts).toContain(
      "\u2610 If this box is checked, the appendix is not used.",
    );
    expect(texts).toContain("Agreed: \u2612");
    // A box that became a plain character is not what the plan placed.
    const flattened = document.replace(
      /<w:sdt><w:sdtPr[^>]*><w:id w:val="11"\/>.*?<w:sdtContent>(.*?)<\/w:sdtContent><\/w:sdt>/u,
      "$1",
    );
    expect(flattened).not.toBe(document);
    expect(
      (
        await verifyFor({
          source,
          template,
          output: await replacePart(docx, "word/document.xml", flattened),
          plan,
        })
      ).ok,
    ).toBe(false);
  });
});
