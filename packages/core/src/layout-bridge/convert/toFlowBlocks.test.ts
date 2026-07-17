import { describe, expect, test } from "bun:test";

import { getTextBoxGroupId } from "../../layout-engine/textBoxGroup";
import { schema } from "../../prosemirror/schema";
import { AUTO_PARAGRAPH_SPACING_PX } from "../../utils/units";
import { toFlowBlocks } from "./toFlowBlocks";

describe("toFlowBlocks paragraph formatting", () => {
  test("stamps each top-level paragraph with its section line pitch", () => {
    const doc = schema.node("doc", null, [
      schema.node("paragraph", { _sectionProperties: { docGrid: { type: "lines", linePitch: 360 } } }, [
        schema.text("First section"),
      ]),
      schema.node("paragraph", null, [schema.text("Final section")]),
    ]);

    const blocks = toFlowBlocks(doc, { finalSectionDocumentGridLinePitchTwips: 480 });
    const paragraphs = blocks.filter((block) => block.kind === "paragraph");

    expect(paragraphs.at(0)?.attrs?.documentGridLinePitch).toBe(24);
    expect(paragraphs.at(1)?.attrs?.documentGridLinePitch).toBe(32);
  });

  test("does not apply the body line grid to table-cell paragraphs", () => {
    const table = schema.node("table", null, [
      schema.node("tableRow", null, [
        schema.node("tableCell", null, [schema.node("paragraph", null, [schema.text("Cell")])]),
      ]),
    ]);
    const blocks = toFlowBlocks(schema.node("doc", null, [table]), {
      finalSectionDocumentGridLinePitchTwips: 360,
    });
    const tableBlock = blocks.at(0);

    expect(tableBlock?.kind).toBe("table");
    if (tableBlock?.kind !== "table") {
      return;
    }
    const cellParagraph = tableBlock.rows.at(0)?.cells.at(0)?.blocks.at(0);
    expect(cellParagraph?.kind).toBe("paragraph");
    if (cellParagraph?.kind !== "paragraph") {
      return;
    }
    expect(cellParagraph.attrs?.documentGridLinePitch).toBeUndefined();
  });

  test("keeps text-box anchors out of paragraph layout", () => {
    const paragraph = toFlowBlocks(
      schema.node("doc", null, [
        schema.node("paragraph", null, [
          schema.text("Before"),
          schema.node("textBoxAnchor", { anchorId: "paragraph:0" }),
          schema.text("After"),
        ]),
      ]),
    ).at(0);
    if (paragraph?.kind !== "paragraph") {
      throw new Error("Expected paragraph");
    }

    expect(paragraph.runs).toHaveLength(2);
    expect(paragraph.runs.flatMap((run) => (run.kind === "text" ? [run.text] : []))).toEqual([
      "Before",
      "After",
    ]);
  });

  test("retains paragraph suppression and stamps the document hyphenation policy", () => {
    const paragraph = toFlowBlocks(
      schema.node("doc", null, [
        schema.node("paragraph", { suppressAutoHyphens: true }, [schema.text("Hyphenation")]),
      ]),
      {
        justificationCompatibility: { type: "legacy" },
        automaticHyphenation: {
          enabled: true,
          doNotHyphenateCaps: true,
          consecutiveLineLimit: 2,
        },
      },
    ).at(0);

    expect(paragraph).toMatchObject({
      kind: "paragraph",
      attrs: {
        suppressAutoHyphens: true,
        justificationCompatibility: { type: "legacy" },
        automaticHyphenation: {
          enabled: true,
          doNotHyphenateCaps: true,
          consecutiveLineLimit: 2,
        },
      },
    });
  });

  test("does not add a default list indent to an authored first-line position", () => {
    const paragraph = toFlowBlocks(
      schema.node("doc", null, [
        schema.node(
          "paragraph",
          {
            numPr: { numId: 1, ilvl: 0 },
            listMarker: "%1)",
            indentFirstLine: 780,
            tabs: [{ position: 1200, alignment: "left" }],
          },
          [schema.text("First item")],
        ),
      ]),
    ).at(0);

    expect(paragraph?.attrs?.indent).toEqual({ firstLine: 52 });
  });

  test("keeps the default indent for a list without authored positioning", () => {
    const paragraph = toFlowBlocks(
      schema.node("doc", null, [
        schema.node("paragraph", { numPr: { numId: 1, ilvl: 1 }, listMarker: "%1.%2." }, [
          schema.text("Nested item"),
        ]),
      ]),
    ).at(0);

    expect(paragraph?.attrs?.indent).toEqual({ left: 96, hanging: 24 });
  });

  test("keeps an explicit list-marker bold override", () => {
    const paragraph = toFlowBlocks(
      schema.node("doc", null, [
        schema.node(
          "paragraph",
          { numPr: { numId: 1, ilvl: 0 }, listMarker: "%1.", listMarkerBold: false },
          [schema.text("List item")],
        ),
      ]),
    ).at(0);

    expect(paragraph?.attrs?.listMarkerBold).toBe(false);
  });

  test("preserves native frame wrap spacing on the positioned container", () => {
    const frame = {
      width: 3600,
      height: 1440,
      hSpace: 144,
      vSpace: 72,
      hAnchor: "page" as const,
      vAnchor: "text" as const,
      x: 720,
      y: 144,
      wrap: "around" as const,
    };
    const blocks = toFlowBlocks(
      schema.node("doc", null, [
        schema.node("paragraph", { _originalFormatting: { frame } }, [schema.text("First")]),
        schema.node("paragraph", { _originalFormatting: { frame } }, [schema.text("Second")]),
        schema.node("paragraph", null, [schema.text("Body")]),
      ]),
    );

    expect(blocks.map((block) => block.kind)).toEqual(["textBox", "paragraph"]);
    expect(blocks.at(0)).toMatchObject({
      kind: "textBox",
      width: 240,
      height: 96,
      displayMode: "float",
      wrapType: "square",
      position: {
        horizontal: { relativeTo: "page", posOffset: 457200 },
        vertical: { relativeTo: "paragraph", posOffset: 91440 },
      },
      content: [{ kind: "paragraph" }, { kind: "paragraph" }],
    });
    const textBox = blocks.at(0);
    expect(textBox?.kind).toBe("textBox");
    if (textBox?.kind !== "textBox") {
      return;
    }
    expect(textBox.distTop).toBeCloseTo(4.8);
    expect(textBox.distBottom).toBeCloseTo(4.8);
    expect(textBox.distLeft).toBeCloseTo(9.6);
    expect(textBox.distRight).toBeCloseTo(9.6);
  });

  test("preserves tables in text box source order and position space", () => {
    const first = schema.node("paragraph", null, [schema.text("Before")]);
    const table = schema.node("table", null, [
      schema.node("tableRow", null, [
        schema.node("tableCell", null, [schema.node("paragraph", null, [schema.text("Cell")])]),
      ]),
    ]);
    const last = schema.node("paragraph", null, [schema.text("After")]);
    const doc = schema.node("doc", null, [
      schema.node("textBox", { width: 240 }, [first, table, last]),
    ]);

    const textBox = toFlowBlocks(doc).at(0);
    if (textBox?.kind !== "textBox") {
      throw new Error("Expected text box block");
    }

    expect(textBox.content.map((block) => block.kind)).toEqual(["paragraph", "table", "paragraph"]);
    expect(textBox.content.map((block) => block.pmStart)).toEqual([
      1,
      1 + first.nodeSize,
      1 + first.nodeSize + table.nodeSize,
    ]);
    expect(textBox.content.at(1)).toMatchObject({
      kind: "table",
      rows: [{ cells: [{ blocks: [{ kind: "paragraph" }] }] }],
    });
  });

  test("keeps drop-cap frames in normal paragraph flow", () => {
    const blocks = toFlowBlocks(
      schema.node("doc", null, [
        schema.node(
          "paragraph",
          { _originalFormatting: { frame: { dropCap: "drop", lines: 3 } } },
          [schema.text("Opening paragraph")],
        ),
      ]),
    );

    expect(blocks.map((block) => block.kind)).toEqual(["paragraph"]);
  });

  test("does not paint an empty structural section-break paragraph", () => {
    const doc = schema.node("doc", null, [
      schema.node("paragraph", { sectionBreakType: "continuous" }),
      schema.node("paragraph", null, [schema.text("Next section")]),
    ]);

    const blocks = toFlowBlocks(doc);

    expect(blocks.map((block) => block.kind)).toEqual(["sectionBreak", "paragraph"]);
  });

  test("paints text in a paragraph that also ends a section", () => {
    const doc = schema.node("doc", null, [
      schema.node("paragraph", { sectionBreakType: "continuous" }, [
        schema.text("Section ending text"),
      ]),
      schema.node("paragraph", null, [schema.text("Next section")]),
    ]);

    const blocks = toFlowBlocks(doc);

    expect(blocks.map((block) => block.kind)).toEqual(["paragraph", "sectionBreak", "paragraph"]);
  });

  test("paints an empty list item that also ends a section", () => {
    const doc = schema.node("doc", null, [
      schema.node("paragraph", {
        sectionBreakType: "continuous",
        numPr: { numId: 1, ilvl: 0 },
        listMarker: "1.",
      }),
      schema.node("paragraph", null, [schema.text("Next section")]),
    ]);

    const blocks = toFlowBlocks(doc);

    expect(blocks.map((block) => block.kind)).toEqual(["paragraph", "sectionBreak", "paragraph"]);
  });

  test("emits a structural break for a standalone column break paragraph", () => {
    const doc = schema.node("doc", null, [
      schema.node("paragraph", null, [schema.text("First column")]),
      schema.node("paragraph", null, [schema.node("hardBreak", { breakType: "column" })]),
      schema.node("paragraph", null, [schema.text("Second column")]),
    ]);

    const blocks = toFlowBlocks(doc);

    expect(blocks.map((block) => block.kind)).toEqual(["paragraph", "columnBreak", "paragraph"]);
    expect(blocks.at(1)).toMatchObject({ kind: "columnBreak", pmStart: 14, pmEnd: 17 });
  });

  test("emits a structural break before text after a leading column break", () => {
    const doc = schema.node("doc", null, [
      schema.node("paragraph", null, [schema.text("First column")]),
      schema.node("paragraph", null, [
        schema.node("hardBreak", { breakType: "column" }),
        schema.text("Second column"),
      ]),
    ]);

    const blocks = toFlowBlocks(doc);

    expect(blocks.map((block) => block.kind)).toEqual(["paragraph", "columnBreak", "paragraph"]);
    expect(blocks.at(1)).toMatchObject({ kind: "columnBreak", pmStart: 15, pmEnd: 16 });
    expect(blocks.at(2)).toMatchObject({
      kind: "paragraph",
      runs: [{ kind: "text", text: "Second column" }],
    });
  });

  test("empty paragraph measurement uses direct paragraph-mark font metrics", () => {
    const doc = schema.node("doc", null, [
      schema.node("paragraph", {
        styleId: "Heading6",
        outlineLevel: 0,
        defaultTextFormatting: { fontSize: 22, fontFamily: { ascii: "Calibri" } },
        _originalFormatting: {
          styleId: "Heading6",
          runProperties: { fontSize: 2, fontFamily: { ascii: "Arial" } },
        },
      }),
    ]);

    const paragraph = toFlowBlocks(doc).at(0);

    expect(paragraph?.kind).toBe("paragraph");
    expect(paragraph?.attrs?.defaultFontSize).toBe(1);
    expect(paragraph?.attrs?.defaultFontFamily).toBe("Arial");
    expect(paragraph?.attrs?.outlineLevel).toBe(0);
    expect(paragraph?.attrs?.reserveEmptyOutlineHeight).toBe(true);
  });

  test("whitespace-only paragraph measurement uses direct paragraph-mark font metrics", () => {
    const doc = schema.node("doc", null, [
      schema.node(
        "paragraph",
        {
          defaultTextFormatting: { fontSize: 22, fontFamily: { ascii: "Calibri" } },
          _originalFormatting: {
            runProperties: { fontSize: 16, fontFamily: { ascii: "Arial Narrow" } },
          },
        },
        [schema.text(" ")],
      ),
    ]);

    const paragraph = toFlowBlocks(doc).at(0);

    expect(paragraph?.kind).toBe("paragraph");
    expect(paragraph?.attrs?.defaultFontSize).toBe(8);
    expect(paragraph?.attrs?.defaultFontFamily).toBe("Arial Narrow");
  });

  test("does not reserve extra outline height away from the start of the story", () => {
    const blocks = toFlowBlocks(
      schema.node("doc", null, [
        schema.node("paragraph", null, [schema.text("content")]),
        schema.node("paragraph", { outlineLevel: 0 }),
      ]),
    );

    expect(blocks.at(1)?.attrs?.outlineLevel).toBe(0);
    expect(blocks.at(1)?.attrs?.reserveEmptyOutlineHeight).toBeUndefined();
  });

  test("ignores a null outline attribute at the layout boundary", () => {
    const paragraph = toFlowBlocks(
      schema.node("doc", null, [schema.node("paragraph", { outlineLevel: null })]),
    ).at(0);

    expect(paragraph?.attrs?.outlineLevel).toBeUndefined();
    expect(paragraph?.attrs?.reserveEmptyOutlineHeight).toBeUndefined();
  });

  test("marks direct formatting on an empty paragraph for spacing layout", () => {
    const doc = schema.node("doc", null, [
      schema.node("paragraph", {
        spaceAfter: 200,
        _originalFormatting: { indentLeft: 720 },
      }),
    ]);

    const paragraph = toFlowBlocks(doc).at(0);

    expect(paragraph?.kind).toBe("paragraph");
    expect(paragraph?.attrs?.hasDirectParagraphFormatting).toBe(true);
  });

  test("does not mark a bare empty paragraph as directly formatted", () => {
    const doc = schema.node("doc", null, [schema.node("paragraph")]);

    const paragraph = toFlowBlocks(doc).at(0);

    expect(paragraph?.kind).toBe("paragraph");
    expect(paragraph?.attrs?.hasDirectParagraphFormatting).toBeUndefined();
  });

  test("tracks paragraph-mark character styling separately from paragraph formatting", () => {
    const doc = schema.node("doc", null, [
      schema.node("paragraph", {
        _originalFormatting: { runProperties: { fontSize: 24 } },
      }),
    ]);

    const paragraph = toFlowBlocks(doc).at(0);

    expect(paragraph?.kind).toBe("paragraph");
    expect(paragraph?.attrs?.hasDirectParagraphFormatting).toBeUndefined();
    expect(paragraph?.attrs?.hasDirectParagraphMarkFormatting).toBe(true);
  });

  test("does not mark an empty paragraph-mark property bag as directly formatted", () => {
    const doc = schema.node("doc", null, [
      schema.node("paragraph", { _originalFormatting: { runProperties: {} } }),
    ]);

    const paragraph = toFlowBlocks(doc).at(0);

    expect(paragraph?.kind).toBe("paragraph");
    expect(paragraph?.attrs?.hasDirectParagraphMarkFormatting).toBeUndefined();
  });

  test("preserves Word rendered-page-break hints for layout", () => {
    const doc = schema.node("doc", null, [
      schema.node("paragraph", { renderedPageBreakBefore: true }, [schema.text("Next page")]),
    ]);

    const paragraph = toFlowBlocks(doc).at(0);

    expect(paragraph?.kind).toBe("paragraph");
    expect(paragraph?.attrs?.renderedPageBreakBefore).toBe(true);
  });

  test("assigns stable block ids for repeated conversions of the same document", () => {
    const doc = schema.node("doc", null, [
      schema.node("paragraph", null, [schema.text("First paragraph")]),
      schema.node("paragraph", null, [schema.text("Second paragraph")]),
    ]);

    const first = toFlowBlocks(doc).map((block) => block.id);
    const second = toFlowBlocks(doc).map((block) => block.id);

    expect(second).toEqual(first);
  });

  test("carries stable PM paraId onto paragraph blocks", () => {
    const doc = schema.node("doc", null, [
      schema.node("paragraph", { paraId: "1A2B3C4D" }, [schema.text("Locate me")]),
    ]);
    const para = toFlowBlocks(doc)[0];
    expect(para?.kind).toBe("paragraph");
    if (para?.kind === "paragraph") {
      expect(para.paraId).toBe("1A2B3C4D");
    }
  });

  test("groups consecutive page-anchored frame paragraphs into one positioned container", () => {
    const frame = {
      width: 1440,
      height: 720,
      hAnchor: "page" as const,
      vAnchor: "page" as const,
      x: 720,
      y: 1440,
      wrap: "around" as const,
    };
    const doc = schema.node("doc", null, [
      schema.node("paragraph", { _originalFormatting: { frame } }, [schema.text("first")]),
      schema.node("paragraph", { _originalFormatting: { frame } }, [schema.text("second")]),
      schema.node("paragraph", null, [schema.text("body")]),
    ]);

    const blocks = toFlowBlocks(doc);

    expect(blocks.map((block) => block.kind)).toEqual(["textBox", "paragraph"]);
    const framed = blocks.at(0);
    if (framed?.kind !== "textBox") {
      throw new Error("Expected grouped paragraph frame");
    }
    expect(framed).toMatchObject({
      width: 96,
      height: 48,
      margins: { top: 0, right: 0, bottom: 0, left: 0 },
      displayMode: "float",
      wrapType: "square",
      position: {
        horizontal: { relativeTo: "page", posOffset: 457_200 },
        vertical: { relativeTo: "page", posOffset: 914_400 },
      },
    });
    expect(
      framed.content.map((contentBlock) =>
        contentBlock.kind === "paragraph" ? contentBlock.runs.at(0) : undefined,
      ),
    ).toMatchObject([
      { kind: "text", text: "first" },
      { kind: "text", text: "second" },
    ]);
  });

  test("keeps page-anchored frames in one wrap set across an empty anchor paragraph", () => {
    const frame = {
      width: 1440,
      height: 720,
      hAnchor: "page" as const,
      vAnchor: "page" as const,
      x: 720,
      y: 1440,
      wrap: "around" as const,
    };
    const doc = schema.node("doc", null, [
      schema.node("paragraph", { _originalFormatting: { frame } }, [schema.text("left")]),
      schema.node("paragraph", { paragraphStyle: "Anchor" }),
      schema.node("paragraph", { _originalFormatting: { frame: { ...frame, x: 2880 } } }, [
        schema.text("right"),
      ]),
      schema.node("paragraph", null, [schema.text("body")]),
    ]);

    const blocks = toFlowBlocks(doc);
    const left = blocks.at(0);
    const right = blocks.at(2);

    expect(blocks.map((block) => block.kind)).toEqual([
      "textBox",
      "paragraph",
      "textBox",
      "paragraph",
    ]);
    expect(left?.kind).toBe("textBox");
    expect(right?.kind).toBe("textBox");
    if (left?.kind === "textBox" && right?.kind === "textBox") {
      expect(getTextBoxGroupId(right)).toBe(getTextBoxGroupId(left));
    }
  });

  test("rejects malformed paragraph attrs at the layout boundary", () => {
    const doc = schema.node("doc", null, [
      schema.node("paragraph", { lineSpacing: "240" }, [schema.text("Invalid paragraph")]),
    ]);

    expect(() => toFlowBlocks(doc)).toThrow("paragraph.attrs.lineSpacing");
  });

  test("rejects malformed mark attrs at the layout boundary", () => {
    const fontSize = schema.mark("fontSize", { size: "large" });
    const doc = schema.node("doc", null, [
      schema.node("paragraph", null, [schema.text("Invalid mark", [fontSize])]),
    ]);

    expect(() => toFlowBlocks(doc)).toThrow("fontSize.attrs.size");
  });

  test("rejects malformed field and math attrs at the layout boundary", () => {
    const fieldDoc = schema.node("doc", null, [
      schema.node("paragraph", null, [
        schema.node("field", {
          fieldType: "NOT_A_FIELD",
          instruction: " PAGE ",
          displayText: "1",
          fieldKind: "simple",
        }),
      ]),
    ]);
    const mathDoc = schema.node("doc", null, [
      schema.node("paragraph", null, [
        schema.node("math", {
          display: "inline",
          ommlXml: 42,
          plainText: "x",
        }),
      ]),
    ]);

    expect(() => toFlowBlocks(fieldDoc)).toThrow("field.attrs.fieldType");
    expect(() => toFlowBlocks(mathDoc)).toThrow("math.attrs.ommlXml");
  });

  test("rejects malformed shape and text box attrs at the layout boundary", () => {
    const shapeDoc = schema.node("doc", null, [
      schema.node("paragraph", null, [schema.node("shape", { width: "wide" })]),
    ]);
    const textBoxDoc = schema.node("doc", null, [
      schema.node("textBox", { width: "wide" }, [
        schema.node("paragraph", null, [schema.text("Invalid text box")]),
      ]),
    ]);

    expect(() => toFlowBlocks(shapeDoc)).toThrow("shape.attrs.width");
    expect(() => toFlowBlocks(textBoxDoc)).toThrow("textBox.attrs.width");
  });

  test("does not convert absent paragraph spacing defaults to zero line height", () => {
    const doc = schema.node("doc", null, [
      schema.node("paragraph", null, [schema.text("First paragraph")]),
    ]);

    const blocks = toFlowBlocks(doc);
    const paragraph = blocks.at(0);

    expect(paragraph?.kind).toBe("paragraph");
    expect(paragraph?.attrs?.spacing).toBeUndefined();
    expect(paragraph?.attrs?.indent).toBeUndefined();
  });

  test("surfaces auto spacing (beforeAutospacing/afterAutospacing) as 14pt for pagination", () => {
    // eigenpal/docx-editor#823: the paged layout reads spacing from the flow
    // block, so auto spacing must override the imported before/after (Word
    // writes `0`) and render the 14pt auto gap while still unedited.
    const doc = schema.node("doc", null, [
      schema.node(
        "paragraph",
        {
          spaceBefore: 0,
          spaceAfter: 0,
          _originalFormatting: {
            beforeAutospacing: true,
            afterAutospacing: true,
            spaceBefore: 0,
            spaceAfter: 0,
          },
          _autospacingBase: { before: 0, after: 0 },
        },
        [schema.text("auto-spaced")],
      ),
    ]);

    const paragraph = toFlowBlocks(doc).at(0);

    expect(paragraph?.kind).toBe("paragraph");
    expect(paragraph?.attrs?.spacing?.before).toBe(AUTO_PARAGRAPH_SPACING_PX);
    expect(paragraph?.attrs?.spacing?.after).toBe(AUTO_PARAGRAPH_SPACING_PX);
    expect(paragraph?.attrs?.automaticSpacing).toEqual({ before: true, after: true });
  });

  test("an explicit spacing edit overrides imported auto-spacing (#823)", () => {
    // Imported with before=0 + beforeAutospacing; the editor then set 240 twips.
    // The edit must win over the now-stale auto-spacing flag.
    const doc = schema.node("doc", null, [
      schema.node(
        "paragraph",
        {
          spaceBefore: 240, // edited away from the imported 0
          _originalFormatting: {
            beforeAutospacing: true,
            spaceBefore: 0,
          },
          _autospacingBase: { before: 0 },
        },
        [schema.text("edited")],
      ),
    ]);

    const paragraph = toFlowBlocks(doc).at(0);

    expect(paragraph?.attrs?.spacing?.before).toBe(16); // 240 twips, not the 14pt auto gap
    expect(paragraph?.attrs?.automaticSpacing).toBeUndefined();
  });

  test("auto spacing still applies when a style supplies spacing the import lacked (#823)", () => {
    // beforeAutospacing with no DIRECT w:before; a style/default placed 200
    // twips into the PM attr. The untouched paragraph is not an edit, so the
    // auto gap must still win over the inherited value.
    const doc = schema.node("doc", null, [
      schema.node(
        "paragraph",
        {
          spaceBefore: 200, // inherited from a style, not a user edit
          _originalFormatting: { beforeAutospacing: true }, // no direct spaceBefore
          _autospacingBase: { before: 200 },
        },
        [schema.text("inherited")],
      ),
    ]);

    const paragraph = toFlowBlocks(doc).at(0);

    expect(paragraph?.attrs?.spacing?.before).toBe(AUTO_PARAGRAPH_SPACING_PX);
  });

  test("auto spacing applies when the auto flag itself came from a style (#823)", () => {
    const doc = schema.node("doc", null, [
      schema.node(
        "paragraph",
        {
          spaceBefore: 200,
          _originalFormatting: { styleId: "AutoSpacing" },
          _autospacingBase: { before: 200 },
        },
        [schema.text("style auto")],
      ),
    ]);

    const paragraph = toFlowBlocks(doc).at(0);

    expect(paragraph?.attrs?.spacing?.before).toBe(AUTO_PARAGRAPH_SPACING_PX);
  });

  test("an edit overrides auto spacing when the import baseline came from a style (#823)", () => {
    const doc = schema.node("doc", null, [
      schema.node(
        "paragraph",
        {
          spaceBefore: 240,
          _originalFormatting: { beforeAutospacing: true },
          _autospacingBase: { before: 200 },
        },
        [schema.text("edited inherited spacing")],
      ),
    ]);

    const paragraph = toFlowBlocks(doc).at(0);

    expect(paragraph?.attrs?.spacing?.before).toBe(16);
  });

  test("keeps auto spacing on an empty paragraph by marking it explicit (#823)", () => {
    // An imported empty paragraph whose only spacing is auto-spacing must keep
    // the 14pt gap; without `spacingExplicit` the empty-paragraph collapse
    // rule in the layout engine would suppress it.
    const doc = schema.node("doc", null, [
      schema.node(
        "paragraph",
        {
          _originalFormatting: {
            beforeAutospacing: true,
            afterAutospacing: true,
          },
          _autospacingBase: { before: null, after: null },
        },
        [],
      ),
    ]);

    const paragraph = toFlowBlocks(doc).at(0);

    expect(paragraph?.attrs?.spacing?.before).toBe(AUTO_PARAGRAPH_SPACING_PX);
    expect(paragraph?.attrs?.spacingExplicit?.before).toBe(true);
    expect(paragraph?.attrs?.spacingExplicit?.after).toBe(true);
  });

  test("keeps document-default spacing on an empty paragraph", () => {
    const doc = schema.node("doc", null, [
      schema.node(
        "paragraph",
        {
          spaceAfter: 160,
          spacingFromDocDefaults: { after: true },
        },
        [],
      ),
    ]);

    const paragraph = toFlowBlocks(doc).at(0);

    expect(paragraph?.attrs?.spacing?.after).toBe(160 / 15);
    expect(paragraph?.attrs?.spacingExplicit?.after).toBe(true);
  });

  test("keeps implicit default-style spacing on an empty paragraph", () => {
    const doc = schema.node("doc", null, [
      schema.node(
        "paragraph",
        {
          spaceBefore: 120,
          spaceAfter: 160,
          spacingFromImplicitDefaultStyle: { before: true, after: true },
        },
        [],
      ),
    ]);

    const paragraph = toFlowBlocks(doc).at(0);

    expect(paragraph?.attrs?.spacing?.before).toBe(120 / 15);
    expect(paragraph?.attrs?.spacing?.after).toBe(160 / 15);
    expect(paragraph?.attrs?.spacingExplicit).toEqual({ before: true, after: true });
  });

  test("suppresses empty hidden list paragraphs and their markers", () => {
    const doc = schema.node("doc", null, [
      schema.node(
        "paragraph",
        {
          defaultTextFormatting: { hidden: true },
          listMarker: "1.",
          listIsBullet: false,
        },
        [],
      ),
    ]);

    const paragraph = toFlowBlocks(doc).at(0);

    expect(paragraph?.attrs?.suppressEmptyParagraphHeight).toBe(true);
    expect(paragraph?.attrs?.listMarkerHidden).toBe(true);
  });

  test("suppresses a paintless imported page-break carrier", () => {
    const doc = schema.node("doc", null, [
      schema.node("paragraph", { _pageBreakCarrier: true, spaceBefore: 360 }),
    ]);

    const paragraph = toFlowBlocks(doc).at(0);

    expect(paragraph?.attrs?.suppressEmptyParagraphHeight).toBe(true);
  });

  test("keeps authored empty paragraphs and visible break-carrier markers", () => {
    const doc = schema.node("doc", null, [
      schema.node("paragraph", { spaceBefore: 360 }),
      schema.node("paragraph", { _pageBreakCarrier: true, listMarker: "1." }),
    ]);

    const blocks = toFlowBlocks(doc);

    expect(blocks.at(0)?.attrs?.suppressEmptyParagraphHeight).toBeUndefined();
    expect(blocks.at(1)?.attrs?.suppressEmptyParagraphHeight).toBeUndefined();
  });

  test("preserves explicit automatic line spacing", () => {
    const doc = schema.node("doc", null, [
      schema.node("paragraph", { lineSpacing: 240, lineSpacingRule: "auto" }, [
        schema.text("First paragraph"),
      ]),
    ]);

    const blocks = toFlowBlocks(doc);
    const paragraph = blocks.at(0);

    expect(paragraph?.kind).toBe("paragraph");
    expect(paragraph?.attrs?.spacing).toEqual({
      line: 1,
      lineRule: "auto",
      lineUnit: "multiplier",
    });
  });
});

describe("toFlowBlocks field handling", () => {
  test("keeps dynamically-rendered field types distinct in layout runs", () => {
    const doc = schema.node("doc", null, [
      schema.node("paragraph", null, [
        schema.node("field", {
          fieldType: "PAGE",
          instruction: " PAGE ",
          displayText: "1",
          fieldKind: "simple",
        }),
        schema.text(" / "),
        schema.node("field", {
          fieldType: "NUMPAGES",
          instruction: " NUMPAGES ",
          displayText: "5",
          fieldKind: "simple",
        }),
        schema.text(" "),
        schema.node("field", {
          fieldType: "DATE",
          instruction: ' DATE \\@ "d MMMM yyyy" ',
          displayText: "29 April 2026",
          fieldKind: "simple",
        }),
      ]),
    ]);

    const blocks = toFlowBlocks(doc);
    const paragraph = blocks.at(0);

    expect(paragraph?.kind).toBe("paragraph");
    if (paragraph?.kind !== "paragraph") {
      throw new Error("Expected paragraph block");
    }

    expect(paragraph.runs.at(0)).toMatchObject({
      kind: "field",
      fieldType: "PAGE",
      fallback: "1",
    });
    expect(paragraph.runs.at(2)).toMatchObject({
      kind: "field",
      fieldType: "NUMPAGES",
      fallback: "5",
    });
    expect(paragraph.runs.at(4)).toMatchObject({
      kind: "field",
      fieldType: "DATE",
      fallback: "29 April 2026",
    });
  });

  test("preserves cached text for field types that are not recomputed by layout", () => {
    const doc = schema.node("doc", null, [
      schema.node("paragraph", null, [
        schema.node("field", {
          fieldType: "MERGEFIELD",
          instruction: ' MERGEFIELD "Client Name" \\* MERGEFORMAT ',
          displayText: "Acme s.r.o.",
          fieldKind: "simple",
        }),
        schema.text(" "),
        schema.node("field", {
          fieldType: "REF",
          instruction: " REF _Ref123 \\h ",
          displayText: "Clause 4.2",
          fieldKind: "complex",
        }),
      ]),
    ]);

    const blocks = toFlowBlocks(doc);
    const paragraph = blocks.at(0);

    expect(paragraph?.kind).toBe("paragraph");
    if (paragraph?.kind !== "paragraph") {
      throw new Error("Expected paragraph block");
    }

    expect(paragraph.runs.at(0)).toMatchObject({
      kind: "field",
      fieldType: "OTHER",
      fallback: "Acme s.r.o.",
    });
    expect(paragraph.runs.at(2)).toMatchObject({
      kind: "field",
      fieldType: "OTHER",
      fallback: "Clause 4.2",
    });
  });

  // Regression: PAGE field rendered with the painter's default font/colour
  // (eigenpal #575) when the bridge skipped extractRunFormatting for field
  // nodes. Word renders a field result with the result run's own w:rPr, so
  // marks attached to the field node must land on the FieldRun.
  test("propagates field-node character marks to the FieldRun formatting", () => {
    const bold = schema.marks["bold"]?.create();
    // fontSize mark stores half-points (the OOXML <w:sz>) — 28 = 14pt.
    const fontSize = schema.marks["fontSize"]?.create({ size: 28 });
    const textColor = schema.marks["textColor"]?.create({ rgb: "FF0000" });
    if (!bold || !fontSize || !textColor) {
      throw new Error("Expected bold/fontSize/textColor marks in schema");
    }
    const doc = schema.node("doc", null, [
      schema.node("paragraph", null, [
        schema.node(
          "field",
          {
            fieldType: "PAGE",
            instruction: " PAGE ",
            displayText: "1",
            fieldKind: "simple",
          },
          undefined,
          [bold, fontSize, textColor],
        ),
      ]),
    ]);

    const blocks = toFlowBlocks(doc);
    const paragraph = blocks.at(0);
    if (paragraph?.kind !== "paragraph") {
      throw new Error("Expected paragraph block");
    }
    const fieldRun = paragraph.runs.at(0);
    if (fieldRun?.kind !== "field") {
      throw new Error("Expected field run");
    }

    expect(fieldRun.bold).toBe(true);
    expect(fieldRun.fontSize).toBe(14);
    expect(fieldRun.color).toBe("#FF0000");
  });
});

describe("toFlowBlocks TOC hyperlink style strip", () => {
  // Regression eigenpal #566: in TOCx paragraphs, Word renders hyperlinks in
  // the paragraph's own colour (no blue + underline). Without stripping the
  // resolved Hyperlink character-style here, the painter's link fallback
  // applies blue + underline and TOC entries look like web links.
  test("strips resolved color/underline on hyperlink text in a TOC paragraph", () => {
    const linkMark = schema.marks["hyperlink"]?.create({
      href: "#_Toc1",
    });
    const underline = schema.marks["underline"]?.create({ style: "single" });
    const textColor = schema.marks["textColor"]?.create({ rgb: "0563C1" });
    if (!linkMark || !underline || !textColor) {
      throw new Error("Expected hyperlink/underline/textColor marks");
    }

    const doc = schema.node("doc", null, [
      schema.node("paragraph", { styleId: "TOC1" }, [
        schema.text("Section 1", [linkMark, underline, textColor]),
      ]),
    ]);

    const blocks = toFlowBlocks(doc);
    const paragraph = blocks.at(0);
    if (paragraph?.kind !== "paragraph") {
      throw new Error("Expected paragraph block");
    }
    const run = paragraph.runs.at(0);
    if (run?.kind !== "text") {
      throw new Error("Expected text run");
    }

    expect(run.hyperlink?.href).toBe("#_Toc1");
    expect(run.hyperlink?.noDefaultStyle).toBe(true);
    expect(run.color).toBeUndefined();
    expect(run.underline).toBeUndefined();
  });

  // The page-number end of a TOC entry is a PAGEREF field inside the
  // hyperlink — the strip must reach field runs too, not just text runs.
  test("strips resolved color/underline on a field run inside a TOC paragraph", () => {
    const linkMark = schema.marks["hyperlink"]?.create({
      href: "#_Toc1",
    });
    const underline = schema.marks["underline"]?.create({ style: "single" });
    const textColor = schema.marks["textColor"]?.create({ rgb: "0563C1" });
    if (!linkMark || !underline || !textColor) {
      throw new Error("Expected hyperlink/underline/textColor marks");
    }

    const doc = schema.node("doc", null, [
      schema.node("paragraph", { styleId: "TOC2" }, [
        schema.node(
          "field",
          {
            fieldType: "PAGEREF",
            instruction: " PAGEREF _Toc1 \\h ",
            displayText: "5",
            fieldKind: "complex",
          },
          undefined,
          [linkMark, underline, textColor],
        ),
      ]),
    ]);

    const blocks = toFlowBlocks(doc);
    const paragraph = blocks.at(0);
    if (paragraph?.kind !== "paragraph") {
      throw new Error("Expected paragraph block");
    }
    const fieldRun = paragraph.runs.at(0);
    if (fieldRun?.kind !== "field") {
      throw new Error("Expected field run");
    }

    expect(fieldRun.hyperlink?.noDefaultStyle).toBe(true);
    expect(fieldRun.color).toBeUndefined();
    expect(fieldRun.underline).toBeUndefined();
  });

  // Non-TOC paragraphs must NOT be stripped — Word still renders normal-body
  // hyperlinks with the Hyperlink character style (blue + underline). The
  // strip is keyed to styleId /^TOC\d*$/i; everything else passes through.
  test("does not strip hyperlinks in non-TOC paragraphs", () => {
    const linkMark = schema.marks["hyperlink"]?.create({
      href: "https://example.com",
    });
    const underline = schema.marks["underline"]?.create({ style: "single" });
    const textColor = schema.marks["textColor"]?.create({ rgb: "0563C1" });
    if (!linkMark || !underline || !textColor) {
      throw new Error("Expected hyperlink/underline/textColor marks");
    }

    const doc = schema.node("doc", null, [
      schema.node("paragraph", null, [schema.text("body link", [linkMark, underline, textColor])]),
    ]);

    const blocks = toFlowBlocks(doc);
    const paragraph = blocks.at(0);
    if (paragraph?.kind !== "paragraph") {
      throw new Error("Expected paragraph block");
    }
    const run = paragraph.runs.at(0);
    if (run?.kind !== "text") {
      throw new Error("Expected text run");
    }

    expect(run.hyperlink?.noDefaultStyle).toBeUndefined();
    expect(run.color).toBeDefined();
    expect(run.underline).toBeDefined();
  });

  // TOC, TOC1..TOC9 should all match. TOCHeading (Word's TOC title) is its
  // own styleId and is NOT a TOC entry — no strip.
  test("TOC styleId regex matches TOC and TOC1..N but not TOCHeading", () => {
    const linkMark = schema.marks["hyperlink"]?.create({ href: "#x" });
    if (!linkMark) {
      throw new Error("Expected hyperlink mark");
    }
    const docFor = (styleId: string) =>
      schema.node("doc", null, [
        schema.node("paragraph", { styleId }, [schema.text("x", [linkMark])]),
      ]);
    const firstRunHyperlinkStripped = (styleId: string) => {
      const blocks = toFlowBlocks(docFor(styleId));
      const para = blocks.at(0);
      if (para?.kind !== "paragraph") {
        throw new Error("Expected paragraph block");
      }
      const run = para.runs.at(0);
      if (run?.kind !== "text") {
        throw new Error("Expected text run");
      }
      return run.hyperlink?.noDefaultStyle === true;
    };

    expect(firstRunHyperlinkStripped("TOC")).toBe(true);
    expect(firstRunHyperlinkStripped("TOC1")).toBe(true);
    expect(firstRunHyperlinkStripped("toc3")).toBe(true);
    expect(firstRunHyperlinkStripped("TOCHeading")).toBe(false);
    expect(firstRunHyperlinkStripped("Normal")).toBe(false);
  });
});

describe("toFlowBlocks table cell formatting", () => {
  test("keeps explicit zero cell margins instead of restoring table defaults", () => {
    const doc = schema.node("doc", null, [
      schema.node("table", { cellMargins: { left: 180, right: 180 } }, [
        schema.node("tableRow", null, [
          schema.node("tableCell", { margins: { left: 0, right: 0 } }, [schema.node("paragraph")]),
        ]),
      ]),
    ]);

    const table = toFlowBlocks(doc).at(0);
    if (table?.kind !== "table") {
      throw new Error("Expected table block");
    }

    expect(table.rows.at(0)?.cells.at(0)?.padding).toMatchObject({ left: 0, right: 0 });
  });

  test("preserves a zero-size styled cell border as a layout-free hairline", () => {
    const hairline = { style: "single", size: 0, color: { rgb: "000000" } };
    const doc = schema.node("doc", null, [
      schema.node("table", null, [
        schema.node("tableRow", null, [
          schema.node("tableCell", { borders: { top: hairline } }, [
            schema.node("paragraph", null, [schema.text("content")]),
          ]),
        ]),
      ]),
    ]);

    const table = toFlowBlocks(doc).at(0);
    if (table?.kind !== "table") {
      throw new Error("Expected table block");
    }

    expect(table.rows.at(0)?.cells.at(0)?.borders?.top).toMatchObject({
      width: 0,
      style: "solid",
    });
  });

  test("carries cantSplit row formatting into layout blocks", () => {
    const doc = schema.node("doc", null, [
      schema.node("table", null, [
        schema.node("tableRow", { _originalFormatting: { cantSplit: true } }, [
          schema.node("tableCell", null, [schema.node("paragraph")]),
        ]),
      ]),
    ]);

    const table = toFlowBlocks(doc).at(0);
    if (table?.kind !== "table") {
      throw new Error("Expected table block");
    }

    expect(table.rows.at(0)?.cantSplit).toBe(true);
  });

  // Regression eigenpal #424 gap 14: the parser captured w:noWrap and the PM
  // schema carried it, but convertTableCell dropped the field, so cells like
  // case numbers / citations wrapped where Word kept them on one line.
  test("threads the cell noWrap attribute into the engine TableCell", () => {
    const doc = schema.node("doc", null, [
      schema.node("table", null, [
        schema.node("tableRow", null, [
          schema.node("tableCell", { noWrap: true }, [
            schema.node("paragraph", null, [schema.text("CASE 123-456")]),
          ]),
          schema.node("tableCell", null, [
            schema.node("paragraph", null, [schema.text("default")]),
          ]),
        ]),
      ]),
    ]);

    const blocks = toFlowBlocks(doc);
    const table = blocks.at(0);
    if (table?.kind !== "table") {
      throw new Error("Expected table block");
    }
    const row = table.rows.at(0);
    if (!row) {
      throw new Error("Expected table row");
    }

    expect(row.cells.at(0)?.noWrap).toBe(true);
    expect(row.cells.at(1)?.noWrap).toBeUndefined();
  });

  test("suppresses an empty terminal paragraph when the cell marker is hidden", () => {
    const doc = schema.node("doc", null, [
      schema.node("table", null, [
        schema.node("tableRow", null, [
          schema.node("tableCell", { hideMark: true }, [schema.node("paragraph")]),
          schema.node("tableCell", null, [schema.node("paragraph")]),
        ]),
      ]),
    ]);

    const table = toFlowBlocks(doc).at(0);
    if (table?.kind !== "table") {
      throw new Error("Expected table block");
    }
    const hiddenMarkerParagraph = table.rows.at(0)?.cells.at(0)?.blocks.at(0);
    const ordinaryParagraph = table.rows.at(0)?.cells.at(1)?.blocks.at(0);

    expect(hiddenMarkerParagraph?.kind).toBe("paragraph");
    expect(ordinaryParagraph?.kind).toBe("paragraph");
    if (hiddenMarkerParagraph?.kind !== "paragraph" || ordinaryParagraph?.kind !== "paragraph") {
      return;
    }
    expect(hiddenMarkerParagraph.attrs?.suppressEmptyParagraphHeight).toBe(true);
    expect(ordinaryParagraph.attrs?.suppressEmptyParagraphHeight).toBeUndefined();
  });

  test("threads the cell text direction into the engine TableCell", () => {
    const doc = schema.node("doc", null, [
      schema.node("table", null, [
        schema.node("tableRow", null, [
          schema.node("tableCell", { textDirection: "btLr" }, [
            schema.node("paragraph", null, [schema.text("Rotated cell")]),
          ]),
          schema.node("tableCell", null, [schema.node("paragraph")]),
        ]),
      ]),
    ]);

    const table = toFlowBlocks(doc).at(0);
    if (table?.kind !== "table") {
      throw new Error("Expected table block");
    }

    expect(table.rows.at(0)?.cells.at(0)?.textDirection).toBe("btLr");
    expect(table.rows.at(0)?.cells.at(1)?.textDirection).toBeUndefined();
  });

  test("keeps an authored trailing empty paragraph after prose cell content", () => {
    const doc = schema.node("doc", null, [
      schema.node("table", null, [
        schema.node("tableRow", null, [
          schema.node("tableCell", null, [
            schema.node("paragraph", null, [schema.text("content")]),
            schema.node("paragraph"),
          ]),
          schema.node("tableCell", null, [schema.node("paragraph")]),
        ]),
      ]),
    ]);

    const table = toFlowBlocks(doc).at(0);
    if (table?.kind !== "table") {
      throw new Error("Expected table block");
    }

    const cells = table.rows.at(0)?.cells;
    expect(cells?.at(0)?.blocks.at(-1)?.attrs?.suppressEmptyParagraphHeight).toBeUndefined();
    expect(cells?.at(1)?.blocks.at(0)?.attrs?.suppressEmptyParagraphHeight).toBeUndefined();
  });

  test("suppresses the required trailing paragraph after a nested table", () => {
    const nestedTable = schema.node("table", null, [
      schema.node("tableRow", null, [
        schema.node("tableCell", null, [schema.node("paragraph", null, [schema.text("nested")])]),
      ]),
    ]);
    const doc = schema.node("doc", null, [
      schema.node("table", null, [
        schema.node("tableRow", null, [
          schema.node("tableCell", null, [nestedTable, schema.node("paragraph")]),
        ]),
      ]),
    ]);

    const table = toFlowBlocks(doc).at(0);
    if (table?.kind !== "table") {
      throw new Error("Expected table block");
    }

    expect(table.rows.at(0)?.cells.at(0)?.blocks.at(-1)?.attrs?.suppressEmptyParagraphHeight).toBe(
      true,
    );
  });

  test("suppresses a terminal run of empty body paragraphs after a final table", () => {
    const doc = schema.node("doc", null, [
      schema.node("table", null, [
        schema.node("tableRow", null, [
          schema.node("tableCell", null, [
            schema.node("paragraph", null, [schema.text("content")]),
          ]),
        ]),
      ]),
      schema.node("paragraph", { tabs: [{ position: 360, alignment: "left" }] }),
      schema.node("paragraph", {
        _originalFormatting: { runProperties: { italic: true } },
      }),
    ]);

    const blocks = toFlowBlocks(doc);

    expect(blocks.map((block) => block.kind)).toEqual(["table", "paragraph", "paragraph"]);
    expect(
      blocks
        .slice(1)
        .map((block) =>
          block.kind === "paragraph" ? block.attrs?.suppressEmptyParagraphHeight : undefined,
        ),
    ).toEqual([true, true]);
  });

  test("keeps terminal empty paragraph height when the preceding block is prose", () => {
    const blocks = toFlowBlocks(
      schema.node("doc", null, [
        schema.node("paragraph", null, [schema.text("content")]),
        schema.node("paragraph"),
      ]),
    );

    expect(blocks.at(-1)?.attrs?.suppressEmptyParagraphHeight).toBeUndefined();
  });

  test("keeps one line before the final marker in a repeated terminal empty run", () => {
    const blocks = toFlowBlocks(
      schema.node("doc", null, [
        schema.node("paragraph", null, [schema.text("content")]),
        schema.node("paragraph"),
        schema.node("paragraph"),
      ]),
    );

    expect(
      blocks
        .slice(-2)
        .map((block) =>
          block.kind === "paragraph" ? block.attrs?.suppressEmptyParagraphHeight : undefined,
        ),
    ).toEqual([undefined, true]);
  });

  test("keeps terminal empty paragraph height when the document contains only empty paragraphs", () => {
    const blocks = toFlowBlocks(
      schema.node("doc", null, [schema.node("paragraph"), schema.node("paragraph")]),
    );

    expect(blocks.map((block) => block.attrs?.suppressEmptyParagraphHeight)).toEqual([
      undefined,
      undefined,
    ]);
  });

  test("keeps a visible empty list item after a final table", () => {
    const blocks = toFlowBlocks(
      schema.node("doc", null, [
        schema.node("table", null, [
          schema.node("tableRow", null, [
            schema.node("tableCell", null, [schema.node("paragraph")]),
          ]),
        ]),
        schema.node("paragraph", {
          numPr: { numId: 1, ilvl: 0 },
          listMarker: "1.",
        }),
      ]),
    );

    expect(blocks.at(-1)?.attrs?.suppressEmptyParagraphHeight).toBeUndefined();
  });

  test("keeps empty paragraphs between a table and later body content", () => {
    const blocks = toFlowBlocks(
      schema.node("doc", null, [
        schema.node("table", null, [
          schema.node("tableRow", null, [
            schema.node("tableCell", null, [schema.node("paragraph")]),
          ]),
        ]),
        schema.node("paragraph"),
        schema.node("paragraph", null, [schema.text("later content")]),
      ]),
    );

    expect(blocks.at(1)?.attrs?.suppressEmptyParagraphHeight).toBeUndefined();
  });
});

describe("toFlowBlocks list numbering", () => {
  test("normalizes Symbol-family bullet markers during flow conversion", () => {
    const doc = schema.node("doc", null, [
      schema.node(
        "paragraph",
        {
          numPr: { numId: 8, ilvl: 0 },
          listIsBullet: true,
          listMarker: "\u00b7",
        },
        [schema.text("Bullet")],
      ),
    ]);

    const blocks = toFlowBlocks(doc);

    expect(blocks.at(0)?.kind).toBe("paragraph");
    expect(blocks.at(0)?.attrs?.listMarker).toBe("\u2022");
  });

  test("marks newly added numbering as a tracked insertion", () => {
    const doc = schema.node("doc", null, [
      schema.node(
        "paragraph",
        {
          numPr: { numId: 1, ilvl: 0 },
          listMarker: "%1.",
          _propertyChanges: [
            {
              type: "paragraphPropertyChange",
              info: { id: 12, author: "Reviewer", date: "2026-01-01" },
              previousFormatting: { numPr: null },
            },
          ],
        },
        [schema.text("Inserted list item")],
      ),
    ]);

    const blocks = toFlowBlocks(doc);

    expect(blocks.at(0)?.attrs?.listMarkerRevision).toEqual({
      kind: "ins",
      author: "Reviewer",
      date: "2026-01-01",
      revisionId: 12,
    });
  });

  test("renders removed numbering as a tracked deletion marker", () => {
    const doc = schema.node("doc", null, [
      schema.node(
        "paragraph",
        {
          _propertyChanges: [
            {
              type: "paragraphPropertyChange",
              info: { id: 13, author: "Reviewer", date: "2026-01-02" },
              previousFormatting: {
                numPr: { numId: 1, ilvl: 0 },
                listIsBullet: false,
                listNumFmt: "decimal",
                listMarker: "%1.",
              },
            },
          ],
        },
        [schema.text("Removed list item")],
      ),
    ]);

    const blocks = toFlowBlocks(doc);

    expect(blocks.at(0)?.attrs?.listMarker).toBe("1.");
    expect(blocks.at(0)?.attrs?.listMarkerRevision).toEqual({
      kind: "del",
      author: "Reviewer",
      date: "2026-01-02",
      revisionId: 13,
    });
  });

  test("removed-numbering deletion numbers off the original stream", () => {
    // An inserted list item (numId 6) advances the final stream to (1). The
    // next item shares that numId but had its numbering removed; in the
    // pre-revision document the inserted item did not exist, so the deleted
    // marker restarts at 1 rather than continuing to 2 off the insertion.
    const doc = schema.node("doc", null, [
      schema.node(
        "paragraph",
        {
          numPr: { numId: 6, ilvl: 0 },
          listMarker: "%1.",
          _propertyChanges: [
            {
              type: "paragraphPropertyChange",
              info: { id: 21, author: "Reviewer", date: "2026-01-01" },
              previousFormatting: { numPr: null },
            },
          ],
        },
        [schema.text("Inserted list item")],
      ),
      schema.node(
        "paragraph",
        {
          _propertyChanges: [
            {
              type: "paragraphPropertyChange",
              info: { id: 22, author: "Reviewer", date: "2026-01-02" },
              previousFormatting: {
                numPr: { numId: 6, ilvl: 0 },
                listIsBullet: false,
                listNumFmt: "decimal",
                listMarker: "%1.",
              },
            },
          ],
        },
        [schema.text("Numbering removed")],
      ),
    ]);

    const blocks = toFlowBlocks(doc);
    const inserted = blocks.at(0);
    const removed = blocks.at(1);
    if (inserted?.kind !== "paragraph" || removed?.kind !== "paragraph") {
      throw new Error("Expected paragraph blocks");
    }

    expect(inserted.attrs?.listMarkerRevision?.kind).toBe("ins");
    expect(removed.attrs?.listMarker).toBe("1.");
    expect(removed.attrs?.listMarkerRevision?.kind).toBe("del");
  });

  test("consecutive removed-numbering deletions advance the original stream", () => {
    // Two adjacent items (numId 6) whose numbering was removed keep their
    // original 1, 2 ordering: the deletion stream advances between them.
    const removedItem = (id: number, text: string) =>
      schema.node(
        "paragraph",
        {
          _propertyChanges: [
            {
              type: "paragraphPropertyChange",
              info: { id, author: "Reviewer", date: "2026-01-02" },
              previousFormatting: {
                numPr: { numId: 6, ilvl: 0 },
                listIsBullet: false,
                listNumFmt: "decimal",
                listMarker: "%1.",
              },
            },
          ],
        },
        [schema.text(text)],
      );
    const doc = schema.node("doc", null, [removedItem(31, "first"), removedItem(32, "second")]);

    const blocks = toFlowBlocks(doc);
    const first = blocks.at(0);
    const second = blocks.at(1);
    if (first?.kind !== "paragraph" || second?.kind !== "paragraph") {
      throw new Error("Expected paragraph blocks");
    }

    expect(first.attrs?.listMarker).toBe("1.");
    expect(second.attrs?.listMarker).toBe("2.");
  });

  test("changed numbering advances the original stream for the previous numId", () => {
    // Item 2's numbering changed from numId 7 to 8 (tracked) — shown as an
    // insertion of the new numId, but in the original document it sat under
    // numId 7. A later numId-7 deletion must continue from that consumed
    // count: a=1, [7->8 changed], deleted=3 — not 2.
    const renumbered = {
      type: "paragraphPropertyChange",
      info: { id: 41, author: "Reviewer", date: "2026-01-01" },
      previousFormatting: {
        numPr: { numId: 7, ilvl: 0 },
        listIsBullet: false,
        listNumFmt: "decimal",
        listMarker: "%1.",
      },
    };
    const doc = schema.node("doc", null, [
      schema.node("paragraph", { numPr: { numId: 7, ilvl: 0 }, listMarker: "%1." }, [
        schema.text("a"),
      ]),
      schema.node(
        "paragraph",
        {
          numPr: { numId: 8, ilvl: 0 },
          listMarker: "%1.",
          _propertyChanges: [renumbered],
        },
        [schema.text("renumbered")],
      ),
      schema.node(
        "paragraph",
        {
          numPr: { numId: 7, ilvl: 0 },
          listMarker: "%1.",
          pPrMark: {
            kind: "del",
            info: { id: 42, author: "Reviewer", date: "2026-01-02" },
          },
        },
        [schema.text("deleted")],
      ),
    ]);

    const blocks = toFlowBlocks(doc);
    const deleted = blocks.at(2);
    if (deleted?.kind !== "paragraph") {
      throw new Error("Expected paragraph block");
    }

    expect(deleted.attrs?.listMarker).toBe("3.");
    expect(deleted.attrs?.listMarkerRevision?.kind).toBe("del");
  });

  test("marks changed numbering as a tracked insertion marker", () => {
    const doc = schema.node("doc", null, [
      schema.node(
        "paragraph",
        {
          numPr: { numId: 1, ilvl: 0 },
          listIsBullet: true,
          listMarker: "\u00b7",
          _propertyChanges: [
            {
              type: "paragraphPropertyChange",
              info: { id: 14, author: "Reviewer", date: "2026-01-03" },
              previousFormatting: {
                numPr: { numId: 2, ilvl: 0 },
                listIsBullet: false,
                listNumFmt: "decimal",
                listMarker: "%1.",
              },
            },
          ],
        },
        [schema.text("Changed list item")],
      ),
    ]);

    const blocks = toFlowBlocks(doc);

    expect(blocks.at(0)?.attrs?.listMarker).toBe("\u2022");
    expect(blocks.at(0)?.attrs?.listMarkerRevision).toEqual({
      kind: "ins",
      author: "Reviewer",
      date: "2026-01-03",
      revisionId: 14,
    });
  });

  test("does not mark unrelated paragraph property changes as list insertions", () => {
    const doc = schema.node("doc", null, [
      schema.node(
        "paragraph",
        {
          numPr: { numId: 1, ilvl: 0 },
          listMarker: "%1.",
          _propertyChanges: [
            {
              type: "paragraphPropertyChange",
              info: { id: 12, author: "Reviewer", date: "2026-01-01" },
              previousFormatting: { alignment: "left" },
            },
          ],
        },
        [schema.text("Plain list item")],
      ),
    ]);

    const blocks = toFlowBlocks(doc);

    expect(blocks.at(0)?.attrs?.listMarkerRevision).toBeUndefined();
  });

  test("formats numbered markers using the paragraph number format", () => {
    const doc = schema.node("doc", null, [
      schema.node(
        "paragraph",
        {
          numPr: { numId: 1, ilvl: 0 },
          listMarker: "%1.",
          listNumFmt: "upperRoman",
        },
        [schema.text("First")],
      ),
      schema.node(
        "paragraph",
        {
          numPr: { numId: 1, ilvl: 0 },
          listMarker: "%1.",
          listNumFmt: "upperRoman",
        },
        [schema.text("Second")],
      ),
    ]);

    const blocks = toFlowBlocks(doc);

    expect(blocks.at(0)?.kind).toBe("paragraph");
    expect(blocks.at(0)?.attrs?.listMarker).toBe("I.");
    expect(blocks.at(1)?.kind).toBe("paragraph");
    expect(blocks.at(1)?.attrs?.listMarker).toBe("II.");
  });

  test("renders repeated placeholders and repeated-letter counters after z", () => {
    const paragraphs = Array.from({ length: 28 }, (_unused, index) =>
      schema.node(
        "paragraph",
        {
          numPr: { numId: 9, ilvl: 0 },
          listMarker: "%1.%1",
          listNumFmt: "lowerLetter",
          listLevelNumFmts: ["lowerLetter"],
        },
        [schema.text(`Item ${index + 1}`)],
      ),
    );

    const blocks = toFlowBlocks(schema.node("doc", null, paragraphs));

    expect(blocks.at(0)?.attrs?.listMarker).toBe("a.a");
    expect(blocks.at(25)?.attrs?.listMarker).toBe("z.z");
    expect(blocks.at(26)?.attrs?.listMarker).toBe("aa.aa");
    expect(blocks.at(27)?.attrs?.listMarker).toBe("bb.bb");
  });

  test("drops unresolved child placeholders with their following punctuation", () => {
    const doc = schema.node("doc", null, [
      schema.node(
        "paragraph",
        {
          numPr: { numId: 2, ilvl: 0 },
          listMarker: "%1.%2.",
          listNumFmt: "decimal",
        },
        [schema.text("Parent")],
      ),
    ]);

    const blocks = toFlowBlocks(doc);

    expect(blocks.at(0)?.kind).toBe("paragraph");
    expect(blocks.at(0)?.attrs?.listMarker).toBe("1.");
  });

  test("formats each level in a multi-level marker with its own number format", () => {
    const doc = schema.node("doc", null, [
      schema.node(
        "paragraph",
        {
          numPr: { numId: 3, ilvl: 0 },
          listMarker: "%1.",
          listNumFmt: "upperRoman",
          listLevelNumFmts: ["upperRoman"],
        },
        [schema.text("Parent")],
      ),
      schema.node(
        "paragraph",
        {
          numPr: { numId: 3, ilvl: 1 },
          listMarker: "%1.%2)",
          listNumFmt: "lowerLetter",
          listLevelNumFmts: ["upperRoman", "lowerLetter"],
        },
        [schema.text("Child")],
      ),
    ]);

    const blocks = toFlowBlocks(doc);

    expect(blocks.at(0)?.attrs?.listMarker).toBe("I.");
    expect(blocks.at(1)?.attrs?.listMarker).toBe("I.a)");
  });

  test("uses authored starts when a nested list begins without parent paragraphs", () => {
    const doc = schema.node("doc", null, [
      schema.node(
        "paragraph",
        {
          numPr: { numId: 8, ilvl: 1 },
          listMarker: "%1.%2.",
          listNumFmt: "decimal",
          listLevelNumFmts: ["decimal", "decimal"],
          listLevelStarts: [3, 3],
        },
        [schema.text("First visible child")],
      ),
      schema.node(
        "paragraph",
        {
          numPr: { numId: 8, ilvl: 1 },
          listMarker: "%1.%2.",
          listNumFmt: "decimal",
          listLevelNumFmts: ["decimal", "decimal"],
          listLevelStarts: [3, 3],
        },
        [schema.text("Second visible child")],
      ),
    ]);

    const blocks = toFlowBlocks(doc);

    expect(blocks.at(0)?.attrs?.listMarker).toBe("3.3.");
    expect(blocks.at(1)?.attrs?.listMarker).toBe("3.4.");
  });

  test("advances nested lists whose authored start is zero", () => {
    const doc = schema.node("doc", null, [
      schema.node(
        "paragraph",
        {
          numPr: { numId: 4, ilvl: 1 },
          listMarker: "%1.%2.",
          listNumFmt: "decimal",
          listLevelStarts: [0, 0],
          listLevelNumFmts: ["decimal", "decimal"],
        },
        [schema.text("First")],
      ),
      schema.node(
        "paragraph",
        {
          numPr: { numId: 4, ilvl: 1 },
          listMarker: "%1.%2.",
          listNumFmt: "decimal",
          listLevelStarts: [0, 0],
          listLevelNumFmts: ["decimal", "decimal"],
        },
        [schema.text("Second")],
      ),
    ]);

    const blocks = toFlowBlocks(doc);

    expect(blocks.at(0)?.attrs?.listMarker).toBe("0.0.");
    expect(blocks.at(1)?.attrs?.listMarker).toBe("0.1.");
  });

  test("formats legal multilevel markers with decimal parent placeholders", () => {
    const paragraphs = [];
    for (let index = 1; index <= 7; index += 1) {
      paragraphs.push(
        schema.node(
          "paragraph",
          {
            numPr: { numId: 7, ilvl: 0 },
            listMarker: "%1",
            listNumFmt: "lowerLetter",
            listLevelNumFmts: ["lowerLetter"],
          },
          [schema.text(`Level ${index}`)],
        ),
      );
    }
    for (let index = 1; index <= 5; index += 1) {
      paragraphs.push(
        schema.node(
          "paragraph",
          {
            numPr: { numId: 7, ilvl: 1 },
            listMarker: "%1.%2",
            listNumFmt: "lowerLetter",
            listLevelNumFmts: ["lowerLetter", "lowerLetter"],
          },
          [schema.text(`Level 7.${index}`)],
        ),
      );
    }
    paragraphs.push(
      schema.node(
        "paragraph",
        {
          numPr: { numId: 7, ilvl: 2 },
          listIsLegal: true,
          listMarker: "%1.%2.%3",
          listNumFmt: "decimal",
          listLevelNumFmts: ["lowerLetter", "lowerLetter", "decimal"],
        },
        [schema.text("Level 7.5.1")],
      ),
    );

    const blocks = toFlowBlocks(schema.node("doc", null, paragraphs));

    expect(blocks.at(-1)?.attrs?.listMarker).toBe("7.5.1");
  });

  test("continues numbering inside text boxes", () => {
    const textBoxNode = schema.nodes.textBox;
    if (!textBoxNode) {
      throw new Error("Expected textBox node in schema");
    }

    const doc = schema.node("doc", null, [
      schema.node("paragraph", { numPr: { numId: 4, ilvl: 0 }, listMarker: "%1." }, [
        schema.text("Before"),
      ]),
      textBoxNode.create(null, [
        schema.node("paragraph", { numPr: { numId: 4, ilvl: 0 }, listMarker: "%1." }, [
          schema.text("Inside"),
        ]),
      ]),
    ]);

    const blocks = toFlowBlocks(doc);
    const textBox = blocks.at(1);

    expect(blocks.at(0)?.attrs?.listMarker).toBe("1.");
    expect(textBox?.kind).toBe("textBox");
    if (textBox?.kind !== "textBox") {
      throw new Error("Expected textBox block");
    }
    const contentBlock = textBox.content.at(0);
    expect(contentBlock?.kind).toBe("paragraph");
    if (contentBlock?.kind !== "paragraph") {
      throw new Error("Expected paragraph in text box");
    }
    expect(contentBlock.attrs?.listMarker).toBe("2.");
  });

  test("carries anchored text-box position into flow blocks", () => {
    const textBoxNode = schema.nodes.textBox;
    if (!textBoxNode) {
      throw new Error("Expected textBox node in schema");
    }

    const position = {
      horizontal: { relativeTo: "margin", align: "center" },
      vertical: { relativeTo: "page", posOffset: 123_456 },
    } as const;
    const doc = schema.node("doc", null, [
      textBoxNode.create(
        {
          wrapType: "topAndBottom",
          position,
        },
        [schema.node("paragraph", null, [schema.text("Inside")])],
      ),
    ]);

    const textBox = toFlowBlocks(doc).at(0);

    expect(textBox?.kind).toBe("textBox");
    if (textBox?.kind !== "textBox") {
      throw new Error("Expected textBox block");
    }
    expect(textBox.wrapType).toBe("topAndBottom");
    expect(textBox.position).toEqual(position);
  });

  test("substitutes style-inherited marker templates without paragraph numPr", () => {
    const doc = schema.node("doc", null, [
      schema.node("paragraph", { numPr: { numId: 5, ilvl: 0 }, listMarker: "%1." }, [
        schema.text("Numbered"),
      ]),
      schema.node("paragraph", { listMarker: "%1." }, [schema.text("Style inherited")]),
    ]);

    const blocks = toFlowBlocks(doc);

    expect(blocks.at(0)?.attrs?.listMarker).toBe("1.");
    expect(blocks.at(1)?.attrs?.listMarker).toBe("1.");
  });
});

// Regression guard for the eigenpal #424 opacity render pipeline. PR #517
// review (gemini-code-assist) flagged that `attrs.opacity !== undefined` in
// buildImageRun allowed the PM schema's null default to leak into
// ImageRun.opacity (typed `number | undefined`). Inline image attributes also
// carry rounded OOXML dimensions, where zero is a valid subpixel result rather
// than an absent value that should receive the default size.
describe("toFlowBlocks image attribute normalization", () => {
  test("omits inline image nodes without a paintable source", () => {
    const doc = schema.node("doc", null, [
      schema.node("paragraph", null, [
        schema.nodes.image.create({
          src: "",
          width: 100,
          height: 100,
        }),
      ]),
    ]);

    const paragraph = toFlowBlocks(doc).at(0);

    expect(paragraph?.kind).toBe("paragraph");
    if (paragraph?.kind === "paragraph") {
      expect(paragraph.runs).toEqual([]);
    }
  });

  test("preserves zero-sized inline image dimensions", () => {
    const doc = schema.node("doc", null, [
      schema.node("paragraph", null, [
        schema.nodes.image.create({
          src: "media/subpixel-placeholder.png",
          width: 1,
          height: 0,
        }),
      ]),
    ]);

    const paragraph = toFlowBlocks(doc).at(0);
    if (paragraph?.kind !== "paragraph") {
      throw new Error("Expected paragraph block");
    }
    const imageRun = paragraph.runs.find((run) => run.kind === "image");

    expect(imageRun?.height).toBe(0);
  });

  test("preserves zero-sized standalone image dimensions", () => {
    const image = schema.nodes.image.create({
      src: "media/subpixel-placeholder.png",
      width: 1,
      height: 0,
    });
    // The schema currently declares images inline-only, but the layout bridge
    // retains a standalone-image conversion path for structural callers.
    const doc = schema.topNodeType.create(null, [image]);

    const block = toFlowBlocks(doc).at(0);

    expect(block?.kind).toBe("image");
    if (block?.kind === "image") {
      expect(block.height).toBe(0);
    }
  });

  test("drops PM null default for inline image opacity (ImageRun)", () => {
    const doc = schema.node("doc", null, [
      schema.node("paragraph", null, [
        schema.nodes.image.create({
          src: "media/image.png",
          width: 100,
          height: 100,
        }),
      ]),
    ]);

    const blocks = toFlowBlocks(doc);
    const paragraph = blocks.at(0);
    if (paragraph?.kind !== "paragraph") {
      throw new Error("Expected paragraph block");
    }
    const imageRun = paragraph.runs.find((run) => run.kind === "image");
    if (imageRun?.kind !== "image") {
      throw new Error("Expected image run");
    }
    // Critical: must be `undefined`, not `null`. The PM schema default is
    // `null` and the bridge must filter it so downstream consumers (the
    // painter, the floating-image collector) see only valid numbers.
    expect(imageRun.opacity).toBeUndefined();
  });

  test("preserves explicit inline image opacity (ImageRun)", () => {
    const doc = schema.node("doc", null, [
      schema.node("paragraph", null, [
        schema.nodes.image.create({
          src: "media/image.png",
          width: 100,
          height: 100,
          opacity: 0.5,
        }),
      ]),
    ]);

    const blocks = toFlowBlocks(doc);
    const paragraph = blocks.at(0);
    if (paragraph?.kind !== "paragraph") {
      throw new Error("Expected paragraph block");
    }
    const imageRun = paragraph.runs.find((run) => run.kind === "image");
    if (imageRun?.kind !== "image") {
      throw new Error("Expected image run");
    }
    expect(imageRun.opacity).toBe(0.5);
  });

  test("preserves authored table-cell anchor scope on image runs", () => {
    const doc = schema.node("doc", null, [
      schema.node("paragraph", null, [
        schema.nodes.image.create({
          src: "media/image.png",
          width: 100,
          height: 100,
          wrapType: "square",
          displayMode: "float",
          layoutInCell: false,
        }),
      ]),
    ]);

    const paragraph = toFlowBlocks(doc).at(0);
    if (paragraph?.kind !== "paragraph") {
      throw new Error("Expected paragraph block");
    }
    const imageRun = paragraph.runs.find((run) => run.kind === "image");

    expect(imageRun?.layoutInCell).toBe(false);
  });

  test("marks embedded-object previews for exact line-height measurement", () => {
    const doc = schema.node("doc", null, [
      schema.node("paragraph", null, [
        schema.nodes.image.create({
          src: "media/preview.png",
          width: 20,
          height: 18,
          _docxObjectPreview: true,
        }),
      ]),
    ]);

    const paragraph = toFlowBlocks(doc).at(0);
    if (paragraph?.kind !== "paragraph") {
      throw new Error("Expected paragraph block");
    }

    expect(paragraph.runs.at(0)).toMatchObject({
      kind: "image",
      exactLineHeight: true,
    });
  });
});
