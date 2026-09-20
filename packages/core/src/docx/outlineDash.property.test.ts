/**
 * `a:prstDash@val` is `ST_PresetLineDashVal`, DrawingML's own dash vocabulary.
 *
 * It used to resolve through a lookup keyed by lower-cased strings that also
 * held CSS `border-style` and CSS `text-decoration-style` keys, so three
 * vocabularies shared one table: `dash` and `solid` collided across all three,
 * and the eight members no other vocabulary spells (`lgDash`, `dashDot`,
 * `lgDashDot`, `lgDashDotDot`, `sysDash`, `sysDot`, `sysDashDot`,
 * `sysDashDotDot`) plus `dot` had no entry at all and painted as a plain line.
 * Nothing could say which table was short, because the table had no domain.
 *
 * Each vocabulary now has its own table and each table is total over its own
 * union. The properties below hold the two halves of that: every member of the
 * enumeration survives a parse → save → parse on both drawing parsers, and
 * each table's domain is exactly its vocabulary, read from the committed
 * schema graph rather than restated here.
 */

import { readFileSync } from "node:fs";

import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import {
  PARSE_WARNING_CODES,
  PRESET_LINE_DASH_VALS,
  presetLineDashToken,
} from "@stll/docx-core/model";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import {
  CSS_BORDER_STROKE_PATTERNS,
  PRESET_DASH_STROKE_PATTERNS,
  UNDERLINE_STROKE_PATTERNS,
} from "../display-list/build/strokes";
import { UNDERLINE_STYLE_VALUES } from "../types/documentEnumValues";
import { CSS_BORDER_STYLE_VALUES } from "../utils/borderCss";
import { parseDocumentBody } from "./documentParser";
import { createParseWarningCollector } from "./parseContext";
import { serializeRun } from "./serializer/runSerializer";
import { parseShapeFromDrawing } from "./shapeParser";
import { parseTextBox } from "./textBoxParser";
import { findDeep, parseXmlDocument } from "./xmlParser";
import type { Shape } from "../types/document";

const schemaEnum = (namespace: string, name: string): readonly string[] => {
  const graph = JSON.parse(
    readFileSync(
      new URL(
        "../../../../specifications/generated/docx-transitional-schema.gen.json",
        import.meta.url,
      ),
      "utf8",
    ),
  ) as { symbols: { id?: string; enumValues?: string[] }[] };
  const symbol = graph.symbols.find(({ id }) => id === `simpleType:{${namespace}}${name}`);
  if (!symbol?.enumValues) {
    throw new Error(`${name} is missing from the schema graph`);
  }
  return symbol.enumValues;
};

const DRAWINGML = "http://schemas.openxmlformats.org/drawingml/2006/main";
const WORDPROCESSINGML = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const ST_PRESET_LINE_DASH_VALS = schemaEnum(DRAWINGML, "ST_PresetLineDashVal");
const ST_UNDERLINE = schemaEnum(WORDPROCESSINGML, "ST_Underline");

const NS = `xmlns:w="${WORDPROCESSINGML}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="${DRAWINGML}" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"`;

const outlineXml = (dash: string) =>
  `<a:ln w="9525"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:prstDash val="${dash}"/></a:ln>`;

const shapeDrawing = (dash: string) => `<w:drawing ${NS}>
  <wp:inline>
    <wp:extent cx="914400" cy="457200"/>
    <wp:docPr id="9" name="Shape 9"/>
    <a:graphic>
      <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
        <wps:wsp>
          <wps:cNvPr id="9" name="Shape 9"/>
          <wps:spPr>
            <a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
            ${outlineXml(dash)}
          </wps:spPr>
          <wps:bodyPr/>
        </wps:wsp>
      </a:graphicData>
    </a:graphic>
  </wp:inline>
</w:drawing>`;

const textBoxDrawing = (dash: string) => `<w:drawing ${NS}>
  <wp:inline>
    <wp:extent cx="914400" cy="457200"/>
    <wp:docPr id="7" name="Text Box 7"/>
    <a:graphic>
      <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
        <wps:wsp>
          <wps:cNvPr id="7" name="Text Box 7"/>
          <wps:spPr>
            <a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
            ${outlineXml(dash)}
          </wps:spPr>
          <wps:txbx><w:txbxContent><w:p/></w:txbxContent></wps:txbx>
          <wps:bodyPr/>
        </wps:wsp>
      </a:graphicData>
    </a:graphic>
  </wp:inline>
</w:drawing>`;

const drawingOf = (xml: string) => {
  const root = parseXmlDocument(xml);
  const drawing = root === null ? null : findDeep(root, "w", "drawing");
  if (!drawing) {
    throw new Error("fixture did not parse");
  }
  return drawing;
};

/**
 * The outline as the editor hands it back: `rawXml` gone, so the serializer
 * has to write the dash from the model rather than replaying the source. That
 * is the leg an authored dash was lost on, and replaying verbatim would pass
 * this property without exercising it.
 */
const rebuiltOutline = (outline: NonNullable<Shape["outline"]>) => {
  const { rawXml: _replaced, ...rebuilt } = outline;
  return rebuilt;
};

const shapeDashRoundTrip = (dash: string) => {
  const parsed = parseShapeFromDrawing(drawingOf(shapeDrawing(dash)));
  if (!parsed?.outline) {
    throw new Error("shape outline did not parse");
  }
  const saved = serializeRun({
    type: "run",
    content: [{ type: "shape", shape: { ...parsed, outline: rebuiltOutline(parsed.outline) } }],
  });
  return {
    parsed: parsed.outline.dash,
    reopened: parseShapeFromDrawing(drawingOf(saved))?.outline?.dash,
  };
};

const textBoxDashRoundTrip = (dash: string) => {
  const parsed = parseTextBox(drawingOf(textBoxDrawing(dash)));
  if (!parsed?.outline) {
    throw new Error("text box outline did not parse");
  }
  // A text box travels through the model as a `textBox` shape; that is the
  // shape the serializer writes and the parser reads back.
  const saved = serializeRun({
    type: "run",
    content: [
      {
        type: "shape",
        shape: {
          type: "shape",
          shapeType: "textBox",
          size: parsed.size,
          outline: rebuiltOutline(parsed.outline),
          textBody: { content: parsed.content },
        },
      },
    ],
  });
  return {
    parsed: parsed.outline.dash,
    reopened: parseTextBox(drawingOf(saved))?.outline?.dash,
  };
};

describe("a preset line dash is read, painted and written in its own vocabulary", () => {
  test("the generated union is the schema's enumeration", () => {
    expect([...PRESET_LINE_DASH_VALS]).toEqual([...ST_PRESET_LINE_DASH_VALS]);
  });

  test(
    "every member round-trips on a shape outline",
    () => {
      fc.assert(
        fc.property(fc.constantFrom(...ST_PRESET_LINE_DASH_VALS), (dash) => {
          expect(shapeDashRoundTrip(dash)).toEqual({ parsed: dash, reopened: dash });
        }),
        propertyConfig(),
      );
    },
    propertyTestTimeout(15_000),
  );

  test(
    "every member round-trips on a text box outline",
    () => {
      fc.assert(
        fc.property(fc.constantFrom(...ST_PRESET_LINE_DASH_VALS), (dash) => {
          expect(textBoxDashRoundTrip(dash)).toEqual({ parsed: dash, reopened: dash });
        }),
        propertyConfig(),
      );
    },
    propertyTestTimeout(15_000),
  );

  test(
    "a token outside the enumeration survives verbatim, as a token",
    () => {
      fc.assert(
        fc.property(
          fc
            .string({ minLength: 1, maxLength: 24 })
            .filter((value) => /^[A-Za-z][A-Za-z0-9]*$/u.test(value))
            .filter((value) => !ST_PRESET_LINE_DASH_VALS.includes(value)),
          (dash) => {
            const { parsed, reopened } = shapeDashRoundTrip(dash);
            // The tri-state, not a widened `string`: a painter that switches
            // on the union cannot mistake an undeclared token for a member.
            expect(parsed).toEqual({ kind: "unrecognised", raw: dash });
            expect(reopened).toEqual({ kind: "unrecognised", raw: dash });
            expect(presetLineDashToken(parsed ?? "solid")).toBe(dash);
          },
        ),
        propertyConfig(),
      );
    },
    propertyTestTimeout(15_000),
  );

  test("an undeclared token is reported through the parse context", () => {
    const collector = createParseWarningCollector("word/document.xml");
    parseTextBox(drawingOf(textBoxDrawing("sysDashDotDot")), collector.context);
    expect(collector.warnings()).toEqual([]);

    parseTextBox(drawingOf(textBoxDrawing("notADash")), collector.context);
    expect(collector.warnings()).toEqual([
      {
        code: PARSE_WARNING_CODES.outlineDashOutsideEnum,
        location: { part: "word/document.xml", element: "a:prstDash" },
        value: "notADash",
        count: 1,
      },
    ]);
  });

  test("the report reaches a document parse, not only the leaf reader", () => {
    const collector = createParseWarningCollector("word/document.xml");
    const body = `<w:document ${NS}><w:body><w:p><w:r>${textBoxDrawing("notADash").replace(`<w:drawing ${NS}>`, "<w:drawing>")}</w:r></w:p></w:body></w:document>`;
    parseDocumentBody(body, null, null, null, null, null, collector.context);

    expect(collector.warnings().map(({ code, value }) => [code, value])).toEqual([
      [PARSE_WARNING_CODES.outlineDashOutsideEnum, "notADash"],
    ]);
  });
});

describe("each stroke table's domain is exactly its own vocabulary", () => {
  test("the preset-dash table covers ST_PresetLineDashVal and nothing else", () => {
    expect(Object.keys(PRESET_DASH_STROKE_PATTERNS).toSorted()).toEqual(
      [...ST_PRESET_LINE_DASH_VALS].toSorted(),
    );
  });

  test("the CSS border table covers the CSS keywords and nothing else", () => {
    expect(Object.keys(CSS_BORDER_STROKE_PATTERNS).toSorted()).toEqual(
      [...CSS_BORDER_STYLE_VALUES].toSorted(),
    );
  });

  test("the underline table covers ST_Underline and nothing else", () => {
    expect(Object.keys(UNDERLINE_STROKE_PATTERNS).toSorted()).toEqual([...ST_UNDERLINE].toSorted());
    // The model's union is the same enumeration, so the table covers both.
    expect([...UNDERLINE_STYLE_VALUES].toSorted()).toEqual([...ST_UNDERLINE].toSorted());
  });

  test("no vocabulary's member resolves through another's table", () => {
    const cssOnly = CSS_BORDER_STYLE_VALUES.filter(
      (style) => !ST_PRESET_LINE_DASH_VALS.includes(style),
    );
    for (const style of cssOnly) {
      expect(Object.hasOwn(PRESET_DASH_STROKE_PATTERNS, style)).toBe(false);
    }
    const dashOnly = ST_PRESET_LINE_DASH_VALS.filter(
      (dash) => !CSS_BORDER_STYLE_VALUES.some((style) => style === dash),
    );
    for (const dash of dashOnly) {
      expect(Object.hasOwn(CSS_BORDER_STROKE_PATTERNS, dash)).toBe(false);
    }
  });
});
