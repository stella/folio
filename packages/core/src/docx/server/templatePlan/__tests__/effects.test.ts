import { beforeAll, describe, expect, test } from "bun:test";

import { applyTemplatePlan } from "../apply";
import {
  findInvisibleEffects,
  removeInvisibleEffects,
  type TextEffect,
} from "../effects";
import { inspectDocxTemplate } from "../inspect";
import { readRawPart } from "../xml";
import {
  buildMinimalDocx,
  buildTemplate,
  NS,
  p,
  r,
  readPart,
  replacePart,
  XML,
} from "./fixtures";
import { applyFor, planFor, unwrap, verifyFor } from "./planning";

const W14 = 'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';

const COLOUR = '<w14:srgbClr w14:val="000000"/>';
const CAMERA = '<w14:camera w14:prst="orthographicFront"/>';
const RIG =
  '<w14:lightRig w14:rig="threePt" w14:dir="t"><w14:rot w14:lat="0" w14:lon="0" w14:rev="0"/></w14:lightRig>';

/** An effect whose values show nothing, and the same effect showing. */
const EFFECTS: Record<TextEffect, { nothing: string[]; shows: string[] }> = {
  glow: {
    nothing: [`<w14:glow w14:rad="0">${COLOUR}</w14:glow>`],
    shows: [
      `<w14:glow w14:rad="63500">${COLOUR}</w14:glow>`,
      // A radius that is not stated is not known to be zero.
      `<w14:glow>${COLOUR}</w14:glow>`,
    ],
  },
  shadow: {
    nothing: [
      `<w14:shadow w14:blurRad="0" w14:dist="0" w14:dir="0" w14:sx="0" w14:sy="0" w14:kx="0" w14:ky="0" w14:algn="none">${COLOUR}</w14:shadow>`,
    ],
    shows: [
      `<w14:shadow w14:blurRad="38100" w14:dist="19050" w14:dir="2700000" w14:sx="100000" w14:sy="100000" w14:kx="0" w14:ky="0" w14:algn="tl">${COLOUR}</w14:shadow>`,
      // Full size straight under the text: hidden only by opaque glyphs.
      `<w14:shadow w14:blurRad="0" w14:dist="0" w14:dir="0" w14:sx="100000" w14:sy="100000" w14:kx="0" w14:ky="0" w14:algn="none">${COLOUR}</w14:shadow>`,
    ],
  },
  reflection: {
    nothing: [
      '<w14:reflection w14:blurRad="0" w14:stA="0" w14:stPos="0" w14:endA="0" w14:endPos="0" w14:dist="0" w14:dir="0" w14:fadeDir="0" w14:sx="0" w14:sy="0" w14:kx="0" w14:ky="0" w14:algn="none"/>',
    ],
    shows: [
      '<w14:reflection w14:blurRad="6350" w14:stA="55000" w14:stPos="0" w14:endA="300" w14:endPos="45500" w14:dist="0" w14:dir="5400000" w14:fadeDir="5400000" w14:sx="100000" w14:sy="-100000" w14:kx="0" w14:ky="0" w14:algn="bl"/>',
    ],
  },
  textOutline: {
    nothing: [
      '<w14:textOutline w14:w="0" w14:cap="rnd" w14:cmpd="sng" w14:algn="ctr"><w14:noFill/><w14:prstDash w14:val="solid"/><w14:bevel/></w14:textOutline>',
      '<w14:textOutline w14:w="9525"><w14:noFill/></w14:textOutline>',
    ],
    shows: [
      `<w14:textOutline w14:w="9525"><w14:solidFill>${COLOUR}</w14:solidFill></w14:textOutline>`,
      // No width is a hairline, not nothing.
      `<w14:textOutline w14:w="0"><w14:solidFill>${COLOUR}</w14:solidFill></w14:textOutline>`,
    ],
  },
  textFill: {
    nothing: [
      '<w14:textFill><w14:solidFill><w14:srgbClr w14:val="1f3864"/></w14:solidFill></w14:textFill>',
    ],
    shows: [
      '<w14:textFill><w14:solidFill><w14:srgbClr w14:val="FF0000"/></w14:solidFill></w14:textFill>',
      '<w14:textFill><w14:solidFill><w14:srgbClr w14:val="1F3864"><w14:alpha w14:val="50000"/></w14:srgbClr></w14:solidFill></w14:textFill>',
      "<w14:textFill><w14:gradFill><w14:gsLst/></w14:gradFill></w14:textFill>",
    ],
  },
  scene3d: {
    nothing: [
      `<w14:scene3d>${CAMERA}${RIG}</w14:scene3d>`,
      `<w14:scene3d><w14:camera w14:prst="orthographicFront"><w14:rot w14:lat="0" w14:lon="0" w14:rev="0"/></w14:camera>${RIG}</w14:scene3d>`,
    ],
    shows: [
      `<w14:scene3d><w14:camera w14:prst="perspectiveFront"/>${RIG}</w14:scene3d>`,
      `<w14:scene3d><w14:camera w14:prst="orthographicFront"><w14:rot w14:lat="0" w14:lon="1200000" w14:rev="0"/></w14:camera>${RIG}</w14:scene3d>`,
    ],
  },
  props3d: {
    nothing: [
      '<w14:props3d w14:extrusionH="0" w14:contourW="0" w14:prstMaterial="none"/>',
      "<w14:props3d/>",
    ],
    shows: [
      '<w14:props3d w14:extrusionH="57150" w14:contourW="0" w14:prstMaterial="warmMatte"/>',
      '<w14:props3d w14:extrusionH="0" w14:contourW="12700" w14:prstMaterial="none"/>',
      '<w14:props3d w14:extrusionH="0" w14:contourW="0"><w14:bevelT w14:w="38100" w14:h="38100" w14:prst="circle"/></w14:props3d>',
    ],
  },
};

/** A part with one run whose properties are `inner`. */
const runPart = (inner: string) =>
  readRawPart(
    "word/document.xml",
    `${XML}<w:document ${NS} ${W14}><w:body><w:p><w:r><w:rPr><w:b/><w:color w:val="1F3864"/>${inner}<w:lang w:val="en-GB"/></w:rPr><w:t>Text</w:t></w:r></w:p></w:body></w:document>`,
  );

const found = (inner: string): TextEffect[] => {
  const part = runPart(inner);
  return part === null
    ? []
    : findInvisibleEffects(part).map(({ effect }) => effect);
};

describe("text effects that show nothing", () => {
  test.each(Object.entries(EFFECTS))(
    "%s: found when its values show nothing, left alone when they show",
    (effect, { nothing, shows }) => {
      for (const xml of nothing) {
        expect(found(xml)).toEqual([effect]);
      }
      for (const xml of shows) {
        expect(found(xml)).toEqual([]);
      }
    },
  );

  test("a 3-D scene stays while the text beside it has depth", () => {
    const scene = EFFECTS.scene3d.nothing.at(0) ?? "";
    expect(found(`${scene}${EFFECTS.props3d.shows.at(0) ?? ""}`)).toEqual([]);
    expect(found(`${scene}${EFFECTS.props3d.nothing.at(0) ?? ""}`)).toEqual([
      "scene3d",
      "props3d",
    ]);
  });

  test("removal takes the effects and nothing else, and a part without one keeps its bytes", () => {
    const nothing = Object.values(EFFECTS)
      .map((effect) => effect.nothing.at(0) ?? "")
      .join("");
    const part = runPart(nothing);
    const plain = runPart("");
    const untouched = runPart(
      Object.values(EFFECTS)
        .slice(0, 5)
        .map((effect) => effect.shows.at(0) ?? "")
        .join(""),
    );
    if (part === null || plain === null || untouched === null) {
      throw new Error("fixture");
    }

    expect(removeInvisibleEffects(part)?.xml).toBe(plain.xml);
    expect(removeInvisibleEffects(part)?.removed).toHaveLength(7);
    expect(removeInvisibleEffects(untouched)?.xml).toBe(untouched.xml);
    expect(removeInvisibleEffects(plain)?.xml).toBe(plain.xml);
  });

  test("an effect is named with what its run properties belong to", () => {
    const scene = EFFECTS.scene3d.nothing.at(0) ?? "";
    const numbering = readRawPart(
      "word/numbering.xml",
      `${XML}<w:numbering ${NS} ${W14}><w:abstractNum w:abstractNumId="4"><w:lvl w:ilvl="0"><w:rPr>${scene}</w:rPr></w:lvl><w:lvl w:ilvl="2"><w:rPr>${scene}</w:rPr></w:lvl></w:abstractNum></w:numbering>`,
    );
    const styles = readRawPart(
      "word/styles.xml",
      `${XML}<w:styles ${NS} ${W14}><w:docDefaults><w:rPrDefault><w:rPr>${scene}</w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:styleId="Heading"><w:rPr>${scene}</w:rPr></w:style></w:styles>`,
    );
    const text = runPart(scene);
    const owners = [numbering, styles, text].flatMap((part) =>
      part === null ? [] : findInvisibleEffects(part).map(({ owner }) => owner),
    );

    expect(owners).toEqual([
      { kind: "level", abstractNumId: 4, ilvl: 0 },
      { kind: "level", abstractNumId: 4, ilvl: 2 },
      { kind: "defaults" },
      { kind: "style", styleId: "Heading" },
      { kind: "text" },
    ]);
  });
});

describe("a plan's output and the template's invisible effects", () => {
  const SCENE = `${EFFECTS.scene3d.nothing.at(0) ?? ""}${EFFECTS.props3d.nothing.at(0) ?? ""}`;
  const GLOW = EFFECTS.glow.shows.at(0) ?? "";
  let template: Uint8Array;
  let source: Uint8Array;

  const edited = async (
    docx: Uint8Array,
    path: string,
    edit: (xml: string) => string,
  ): Promise<Uint8Array> =>
    await replacePart(docx, path, edit((await readPart(docx, path)) ?? ""));

  beforeAll(async () => {
    source = await buildMinimalDocx(p(r("Body.")));
    // A level whose number carries the default 3-D scene, a style with one
    // and with a glow that shows, and a header run with one.
    let built = await buildTemplate();
    built = await edited(built, "word/numbering.xml", (xml) =>
      xml
        .replace("<w:numbering ", `<w:numbering ${W14} `)
        .replace("</w:pPr></w:lvl>", `</w:pPr><w:rPr>${SCENE}</w:rPr></w:lvl>`),
    );
    built = await edited(built, "word/styles.xml", (xml) =>
      xml
        .replace("<w:styles ", `<w:styles ${W14} `)
        .replace(
          "<w:b/><w:caps/></w:rPr>",
          `<w:b/><w:caps/>${SCENE}${GLOW}</w:rPr>`,
        ),
    );
    template = await edited(built, "word/header1.xml", (xml) =>
      xml
        .replace("<w:hdr ", `<w:hdr ${W14} `)
        .replace("<w:r><w:t", `<w:r><w:rPr>${SCENE}</w:rPr><w:t`),
    );
  });

  test("the inspection lists them with their owners", async () => {
    const { invisibleEffects } = unwrap(await inspectDocxTemplate(template));

    expect(
      invisibleEffects.map(({ part, effect, owner }) => [
        part,
        effect,
        owner.kind,
      ]),
    ).toEqual([
      ["word/numbering.xml", "scene3d", "level"],
      ["word/numbering.xml", "props3d", "level"],
      ["word/styles.xml", "scene3d", "style"],
      ["word/styles.xml", "props3d", "style"],
      ["word/header1.xml", "scene3d", "text"],
      ["word/header1.xml", "props3d", "text"],
    ]);
  });

  test("the output is written without them, and an effect that shows stays", async () => {
    const { docx, report } = unwrap(await applyFor({ source, template }));
    const [numbering, styles, header] = await Promise.all(
      ["word/numbering.xml", "word/styles.xml", "word/header1.xml"].map(
        async (path) => (await readPart(docx, path)) ?? "",
      ),
    );

    expect(report.ok).toBe(true);
    expect(`${numbering}${styles}${header}`).not.toContain("scene3d");
    expect(`${numbering}${styles}${header}`).not.toContain("props3d");
    expect(styles).toContain(GLOW);
    expect(report.textEffects).toEqual({
      removed: [
        { part: "word/numbering.xml", effect: "scene3d", count: 1 },
        { part: "word/numbering.xml", effect: "props3d", count: 1 },
        { part: "word/styles.xml", effect: "scene3d", count: 1 },
        { part: "word/styles.xml", effect: "props3d", count: 1 },
        { part: "word/header1.xml", effect: "scene3d", count: 1 },
        { part: "word/header1.xml", effect: "props3d", count: 1 },
      ],
      left: [],
    });
    // Nothing else of the parts changed.
    expect(numbering).toBe(
      ((await readPart(template, "word/numbering.xml")) ?? "").replace(
        SCENE,
        "",
      ),
    );
  });

  test("a plan that keeps them writes the parts as they are", async () => {
    const plan = await planFor({
      source,
      template,
      extra: { textEffects: "keep" },
    });
    const { docx, report } = unwrap(
      await applyTemplatePlan({ source, template, plan }),
    );

    for (const path of ["word/numbering.xml", "word/styles.xml"]) {
      // oxlint-disable-next-line no-await-in-loop -- two small reads
      expect(await readPart(docx, path)).toBe(await readPart(template, path));
    }
    expect(report.textEffects.removed).toEqual([]);
    expect(report.textEffects.left).toHaveLength(6);
    expect(
      (await verifyFor({ source, template, output: docx, plan })).issues,
    ).toEqual([]);
  });

  test("a template without such effects gives a package with the same parts", async () => {
    const plainTemplate = await buildTemplate();
    const { docx, report } = unwrap(
      await applyFor({ source, template: plainTemplate }),
    );

    expect(report.textEffects).toEqual({ removed: [], left: [] });
    for (const path of ["word/styles.xml", "word/header1.xml"]) {
      // oxlint-disable-next-line no-await-in-loop -- two small reads
      expect(await readPart(docx, path)).toBe(
        await readPart(plainTemplate, path),
      );
    }
  });
});
