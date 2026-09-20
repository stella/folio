/**
 * An explicit off is not an absent attribute.
 *
 * `ST_OnOff` has three states: absent, explicit off (`0|false|off`) and
 * explicit on (`1|true|on`). None of the attributes covered here carries an XSD
 * default, so "written with its default" is not available as an excuse: a
 * serializer that emits the attribute only when the model says `true` writes an
 * explicit off back as an absence, and the document changes.
 *
 * `parseOnOffValue` already keeps all three states, so the model does too. This
 * file is the one place that holds every emit path to that contract, so a new
 * `ST_OnOff` slot gets its three states pinned by adding a row rather than a
 * file.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";
import { mergeParagraphFormatting } from "../utils/paragraphFormattingMerge";

import { parseParagraphProperties } from "./paragraphParser";
import { parseSectionProperties } from "./sectionParser";
import { serializeBorder } from "./serializer/borderSerializer";
import { parseStyles } from "./styleParser";
import { parseTableCellProperties, parseTableProperties } from "./tableParser";
import { parseXmlDocument } from "./xmlParser";

const WORD_NAMESPACE = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

/** Every spelling `ST_OnOff` accepts, plus the absence it must stay distinct from. */
const ON_OFF_SPELLINGS = [undefined, "0", "false", "off", "1", "true", "on"] as const;

type OnOffSpelling = (typeof ON_OFF_SPELLINGS)[number];

/** The state a spelling authors, written out rather than derived from the reader under test. */
const statedBy = (spelling: OnOffSpelling): boolean | undefined => {
  switch (spelling) {
    case undefined:
      return undefined;
    case "0":
    case "false":
    case "off":
      return false;
    case "1":
    case "true":
    case "on":
      return true;
    default:
      return spelling satisfies never;
  }
};

const parseOne = (xml: string) => {
  const element = parseXmlDocument(xml);
  if (!element) {
    throw new Error("fixture did not parse");
  }
  return element;
};

const attribute = (name: string, spelling: OnOffSpelling): string =>
  spelling === undefined ? "" : ` w:${name}="${spelling}"`;

describe("CT_Border @shadow and @frame", () => {
  /** The two toggles `CT_Border` carries, on every border-bearing container. */
  const BORDER_TOGGLES = ["shadow", "frame"] as const;

  /**
   * Each container's real entry point, not the shared helper: the property is
   * that every tier agrees, which only a per-tier read can show. `w:top` is the
   * side every one of them declares.
   */
  const TIERS = {
    paragraph: (border: string) =>
      parseParagraphProperties(
        parseOne(`<w:pPr ${WORD_NAMESPACE}><w:pBdr>${border}</w:pBdr></w:pPr>`),
        null,
      )?.borders?.top,
    style: (border: string) =>
      parseStyles(
        `<w:styles ${WORD_NAMESPACE}>
           <w:style w:type="paragraph" w:styleId="Bordered">
             <w:name w:val="Bordered"/>
             <w:pPr><w:pBdr>${border}</w:pBdr></w:pPr>
           </w:style>
         </w:styles>`,
        null,
      ).get("Bordered")?.pPr?.borders?.top,
    table: (border: string) =>
      parseTableProperties(
        parseOne(`<w:tblPr ${WORD_NAMESPACE}><w:tblBorders>${border}</w:tblBorders></w:tblPr>`),
      )?.borders?.top,
    cell: (border: string) =>
      parseTableCellProperties(
        parseOne(`<w:tcPr ${WORD_NAMESPACE}><w:tcBorders>${border}</w:tcBorders></w:tcPr>`),
      )?.borders?.top,
    page: (border: string) =>
      parseSectionProperties(
        parseOne(`<w:sectPr ${WORD_NAMESPACE}><w:pgBorders>${border}</w:pgBorders></w:sectPr>`),
      ).pageBorders?.top,
  } as const;

  const TIER_NAMES = Object.keys(TIERS) as (keyof typeof TIERS)[];

  const borderXml = (toggle: (typeof BORDER_TOGGLES)[number], spelling: OnOffSpelling): string =>
    `<w:top w:val="single" w:sz="8" w:space="0"${attribute(toggle, spelling)}/>`;

  test(
    "every spelling survives a forced save, in every container",
    () => {
      fc.assert(
        fc.property(
          fc.constantFrom(...TIER_NAMES),
          fc.constantFrom(...BORDER_TOGGLES),
          fc.constantFrom(...ON_OFF_SPELLINGS),
          (tier, toggle, spelling) => {
            const parsed = TIERS[tier](borderXml(toggle, spelling));
            if (!parsed) {
              throw new Error(`border did not parse on the ${tier} tier`);
            }
            expect(parsed[toggle]).toBe(statedBy(spelling));

            const reparsed = TIERS[tier](serializeBorder(parsed, "top"));
            expect(reparsed?.[toggle]).toBe(statedBy(spelling));
          },
        ),
        propertyConfig({ numRuns: 250 }),
      );
    },
    propertyTestTimeout(15_000),
  );

  /**
   * `CT_Border` attributes do not inherit one by one: a tier that states a
   * side replaces the whole side, so the effective toggle is the stating
   * tier's, and an unstated toggle on a stated side is not the style's.
   */
  const effectiveShadow = ({
    direct,
    fromStyle,
  }: {
    direct: OnOffSpelling | "no side";
    fromStyle: OnOffSpelling | "no side";
  }): boolean | undefined => {
    if (direct !== "no side") {
      return statedBy(direct);
    }
    if (fromStyle !== "no side") {
      return statedBy(fromStyle);
    }
    return undefined;
  };

  const SIDE_STATES = ["no side", ...ON_OFF_SPELLINGS] as const;

  test(
    "the effective shadow under an inheriting style survives a forced save",
    () => {
      fc.assert(
        fc.property(
          fc.constantFrom(...SIDE_STATES),
          fc.constantFrom(...SIDE_STATES),
          (direct, fromStyle) => {
            const sideOf = (state: (typeof SIDE_STATES)[number]) =>
              state === "no side" ? "" : borderXml("shadow", state);

            const saveBorders = (borders: { top?: unknown } | undefined) => {
              const top = (borders as { top?: Parameters<typeof serializeBorder>[0] } | undefined)
                ?.top;
              return top === undefined ? undefined : TIERS.paragraph(serializeBorder(top, "top"));
            };

            const styleFormatting = parseStyles(
              `<w:styles ${WORD_NAMESPACE}>
                 <w:style w:type="paragraph" w:styleId="Bordered">
                   <w:name w:val="Bordered"/>
                   <w:pPr><w:pBdr>${sideOf(fromStyle)}</w:pBdr></w:pPr>
                 </w:style>
               </w:styles>`,
              null,
            ).get("Bordered")?.pPr;
            const directFormatting = parseParagraphProperties(
              parseOne(`<w:pPr ${WORD_NAMESPACE}><w:pBdr>${sideOf(direct)}</w:pBdr></w:pPr>`),
              null,
            );

            const saved = {
              ...directFormatting,
              borders:
                saveBorders(directFormatting?.borders) === undefined
                  ? undefined
                  : { top: saveBorders(directFormatting?.borders) },
            };
            const resolved = mergeParagraphFormatting(styleFormatting, saved);

            expect(resolved?.borders?.top?.shadow).toBe(effectiveShadow({ direct, fromStyle }));
          },
        ),
        propertyConfig({ numRuns: 250 }),
      );
    },
    propertyTestTimeout(15_000),
  );
});
