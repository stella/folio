import { Result } from "better-result";
import { beforeAll, describe, expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { applyTemplatePlan } from "../apply";
import { analyzeDocxBlocks } from "../blocks";
import { isManualLabel } from "../plan";
import type {
  TemplatePlan,
  TemplatePlanItem,
  TemplatePlanSourceItem,
} from "../types";
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
            instruction: "\\* MERGEFORMAT",
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
