import { describe, expect, test } from "bun:test";

import { serializeStylesXml } from "./serializer/stylesSerializer";
import { getDefaultParagraphStyle, parseStyleDefinitions, parseStyles } from "./styleParser";

const STYLES_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

describe("docDefaults presence (#909)", () => {
  test("an empty but present docDefaults parses to a defined object", () => {
    // Lets the resolver tell "document declares empty defaults" (zero spacing)
    // apart from "no docDefaults at all", so it does not synthesize the
    // built-in Normal spacing over an explicitly-empty default.
    const defs = parseStyleDefinitions(
      `<w:styles ${STYLES_NS}>
        <w:docDefaults><w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>
      </w:styles>`,
      null,
    );
    expect(defs.docDefaults).toBeDefined();
  });

  test("a document with no docDefaults element leaves docDefaults undefined", () => {
    const defs = parseStyleDefinitions(`<w:styles ${STYLES_NS}></w:styles>`, null);
    expect(defs.docDefaults).toBeUndefined();
  });

  test("preserves document-default language metadata", () => {
    const defs = parseStyleDefinitions(
      `<w:styles ${STYLES_NS}>
        <w:docDefaults>
          <w:rPrDefault>
            <w:rPr><w:lang w:val="cs-CZ" w:eastAsia="ja-JP" w:bidi="ar-SA"/></w:rPr>
          </w:rPrDefault>
        </w:docDefaults>
      </w:styles>`,
      null,
    );

    expect(defs.docDefaults?.rPr?.language).toEqual({
      val: "cs-CZ",
      eastAsia: "ja-JP",
      bidi: "ar-SA",
    });
  });
});

describe("style default ST_OnOff values", () => {
  test("recognizes on as a default paragraph style value", () => {
    const styles = parseStyles(
      `<w:styles ${STYLES_NS}>
        <w:style w:type="paragraph" w:default="on" w:styleId="BodyText">
          <w:name w:val="Body Text"/>
        </w:style>
      </w:styles>`,
      null,
    );

    expect(styles.get("BodyText")?.default).toBe(true);
    expect(getDefaultParagraphStyle(styles)?.styleId).toBe("BodyText");
  });

  test("uses the same decoder for latent and table style flags", () => {
    const definitions = parseStyleDefinitions(
      `<w:styles ${STYLES_NS}>
        <w:latentStyles w:defLockedState="on" w:defQFormat="off"/>
        <w:style w:type="table" w:styleId="Grid">
          <w:tblPr><w:tblLook w:firstRow="on" w:lastRow="off"/></w:tblPr>
        </w:style>
      </w:styles>`,
      null,
    );

    expect(definitions.latentStyles).toMatchObject({
      defLockedState: true,
      defQFormat: false,
    });
    expect(definitions.styles.at(0)?.tblPr?.look).toMatchObject({
      firstRow: true,
      lastRow: false,
    });
  });
});

describe("style run toggle parsing", () => {
  test("the final occurrence within one run-properties level wins", () => {
    const styles = parseStyles(
      `<w:styles ${STYLES_NS}>
        <w:style w:type="character" w:styleId="OffLast">
          <w:rPr><w:b/><w:b w:val="0"/></w:rPr>
        </w:style>
        <w:style w:type="character" w:styleId="OnLast">
          <w:rPr><w:b w:val="false"/><w:b/></w:rPr>
        </w:style>
      </w:styles>`,
      null,
    );

    expect(styles.get("OffLast")?.rPr?.bold).toBe(false);
    expect(styles.get("OnLast")?.rPr?.bold).toBe(true);
  });
});

describe("style horizontal text scale parsing", () => {
  const parseStyleScale = (value: string) =>
    parseStyles(
      `<w:styles ${STYLES_NS}>
        <w:style w:type="character" w:styleId="Scaled">
          <w:rPr><w:w w:val="${value}"/></w:rPr>
        </w:style>
      </w:styles>`,
      null,
    ).get("Scaled")?.rPr?.scale;

  test.each([
    [" 0% ", 0],
    ["600%", 600],
  ])("preserves strict scale %s", (value, expected) => {
    expect(parseStyleScale(value)).toBe(expected);
  });

  test.each(["", "   ", "0garbage", "1e2", "600oops", "0x10", "601%"])(
    "drops malformed scale %p",
    (value) => {
      expect(parseStyleScale(value)).toBeUndefined();
    },
  );
});

describe("style tab leader normalization", () => {
  test("normalizes the default leader to a save/reopen fixed point", () => {
    const parsed = parseStyleDefinitions(
      `<w:styles ${STYLES_NS}>
        <w:docDefaults><w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>
        <w:style w:type="paragraph" w:styleId="BodyText">
          <w:name w:val="Body Text"/>
          <w:pPr>
            <w:tabs>
              <w:tab w:val="left" w:pos="720" w:leader="none"/>
              <w:tab w:val="right" w:pos="1440" w:leader="dot"/>
            </w:tabs>
          </w:pPr>
        </w:style>
      </w:styles>`,
      null,
    );

    expect(parsed.styles.at(0)?.pPr?.tabs).toEqual([
      { position: 720, alignment: "left" },
      { position: 1440, alignment: "right", leader: "dot" },
    ]);

    expect(parseStyleDefinitions(serializeStylesXml(parsed), null)).toEqual(parsed);
  });
});

describe("style table measurements", () => {
  test("defaults missing w:type to dxa", () => {
    const styles = parseStyles(
      `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:style w:type="table" w:styleId="TableGrid">
          <w:name w:val="Table Grid"/>
          <w:tblPr>
            <w:tblW w:w="5000"/>
          </w:tblPr>
        </w:style>
      </w:styles>`,
      null,
    );

    expect(styles.get("TableGrid")?.tblPr?.width).toEqual({
      type: "dxa",
      value: 5000,
    });
  });
});

describe("style inheritance cycles", () => {
  test("basedOn cycle terminates and merges both styles' properties", () => {
    const styles = parseStyles(
      `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:style w:type="character" w:styleId="CycleA">
          <w:name w:val="Cycle A"/>
          <w:basedOn w:val="CycleB"/>
          <w:rPr><w:b/></w:rPr>
        </w:style>
        <w:style w:type="character" w:styleId="CycleB">
          <w:name w:val="Cycle B"/>
          <w:basedOn w:val="CycleA"/>
          <w:rPr><w:i/></w:rPr>
        </w:style>
      </w:styles>`,
      null,
    );

    const cycleA = styles.get("CycleA");
    expect(cycleA?.rPr?.bold).toBe(true);
    expect(cycleA?.rPr?.italic).toBe(true);
    // The cycle guard stops the second visit, so CycleB's own italic must
    // still be present after merging with CycleA.
    const cycleB = styles.get("CycleB");
    expect(cycleB?.rPr?.italic).toBe(true);
  });

  test("self-referential basedOn terminates", () => {
    const styles = parseStyles(
      `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:style w:type="paragraph" w:styleId="Selfish">
          <w:name w:val="Selfish"/>
          <w:basedOn w:val="Selfish"/>
          <w:rPr><w:b/></w:rPr>
        </w:style>
      </w:styles>`,
      null,
    );

    expect(styles.get("Selfish")?.rPr?.bold).toBe(true);
  });
});

describe("style toggle inheritance (ECMA-376 §17.7.3)", () => {
  test("an explicit off in a derived style turns off a bold base style", () => {
    const styles = parseStyles(
      `<w:styles ${STYLES_NS}>
        <w:style w:type="character" w:styleId="BoldBase">
          <w:name w:val="Bold Base"/>
          <w:rPr><w:b/></w:rPr>
        </w:style>
        <w:style w:type="character" w:styleId="NotBold">
          <w:name w:val="Not Bold"/>
          <w:basedOn w:val="BoldBase"/>
          <w:rPr><w:b w:val="0"/></w:rPr>
        </w:style>
      </w:styles>`,
      null,
    );

    expect(styles.get("BoldBase")?.rPr?.bold).toBe(true);
    expect(styles.get("NotBold")?.rPr?.bold).toBe(false);
  });

  test("an explicit on in a derived style wins over a base style's off", () => {
    const styles = parseStyles(
      `<w:styles ${STYLES_NS}>
        <w:style w:type="paragraph" w:styleId="ItalicOffBase">
          <w:name w:val="Italic Off Base"/>
          <w:rPr><w:i w:val="false"/></w:rPr>
        </w:style>
        <w:style w:type="paragraph" w:styleId="ItalicAgain">
          <w:name w:val="Italic Again"/>
          <w:basedOn w:val="ItalicOffBase"/>
          <w:rPr><w:i/></w:rPr>
        </w:style>
      </w:styles>`,
      null,
    );

    expect(styles.get("ItalicOffBase")?.rPr?.italic).toBe(false);
    expect(styles.get("ItalicAgain")?.rPr?.italic).toBe(true);
  });

  test("an intermediate style that never mentions the toggle falls through untouched", () => {
    const styles = parseStyles(
      `<w:styles ${STYLES_NS}>
        <w:style w:type="character" w:styleId="ShadowBase">
          <w:name w:val="Shadow Base"/>
          <w:rPr><w:shadow/></w:rPr>
        </w:style>
        <w:style w:type="character" w:styleId="ShadowMiddle">
          <w:name w:val="Shadow Middle"/>
          <w:basedOn w:val="ShadowBase"/>
          <w:rPr><w:smallCaps/></w:rPr>
        </w:style>
        <w:style w:type="character" w:styleId="ShadowOff">
          <w:name w:val="Shadow Off"/>
          <w:basedOn w:val="ShadowMiddle"/>
          <w:rPr><w:shadow w:val="0"/></w:rPr>
        </w:style>
      </w:styles>`,
      null,
    );

    // The middle style declares no <w:shadow>, so it must keep inheriting the
    // base's "on" state rather than being treated as an implicit "off".
    expect(styles.get("ShadowMiddle")?.rPr?.shadow).toBe(true);
    expect(styles.get("ShadowMiddle")?.rPr?.smallCaps).toBe(true);
    // The leaf's explicit off wins over the base two levels up.
    expect(styles.get("ShadowOff")?.rPr?.shadow).toBe(false);
  });
});

describe("style borders", () => {
  test("keeps a w:val outside ST_Border verbatim", () => {
    const styles = parseStyles(
      `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:style w:type="paragraph" w:styleId="BodyText">
          <w:name w:val="Body Text"/>
          <w:pPr>
            <w:pBdr>
              <w:top w:val="dashDotDot" w:sz="8" w:color="FF0000"/>
            </w:pBdr>
          </w:pPr>
        </w:style>
      </w:styles>`,
      null,
    );

    expect(styles.get("BodyText")?.pPr?.borders?.top).toMatchObject({
      color: { rgb: "FF0000" },
      size: 8,
      style: { kind: "unrecognised", raw: "dashDotDot" },
    });
  });
});
