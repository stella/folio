/**
 * A renamed name reads the same in either spelling, everywhere folio reads it.
 *
 * `check-strict-name-tolerance.ts` catches a reader that hand-lists both
 * spellings. It cannot catch the other half of the class — a reader that takes
 * only the Transitional one — because the thing that decides whether a read
 * needs the tolerance is the receiver's **complex type**, and the source never
 * says it: `findChild(pBdr, "w", "left")` and `findChild(tblBorders, "w",
 * "left")` are the same six tokens, and only the second sits on a type the
 * schema declares the rename for. A name-only rule would fire on `w:pgMar`,
 * `w:pgBorders` and `w:pBdr`, none of which is renamed, and buy its silence
 * with an exemption list. So this test is the guard for that half.
 *
 * The slots come from the generated table, so a rename a schema refresh adds
 * arrives here with no fixture and fails. The **places** folio reads each slot
 * cannot come from the schema — they are facts about folio's own parsers, and
 * the style parser's private copies of the table readers are exactly why this
 * test exists — so they are written out below, held total over the generated
 * table by the compiler. `null` is a decision too: folio models neither
 * spelling of an indent in character units, so there is no model to keep equal.
 */

import { describe, expect, test } from "bun:test";

import { parseNumbering } from "./numberingParser";
import { parseParagraphProperties } from "./paragraphParser";
import { STRICT_NAMES, type StrictName, TRANSITIONAL_NAME_BY_STRICT_NAME } from "./strictNames.gen";
import { parseStylesPackage } from "./styleParser";
import { parseTableCellProperties, parseTableProperties } from "./tableParser";
import { parseXmlDocument } from "./xmlParser";

const WML = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

/** One place folio reads a renamed slot: the markup, in a named spelling, and what folio makes of it. */
type ReadSite = {
  /** The reader under test, for the failure message. */
  site: string;
  /** The fragment or part, with the renamed name spelled as asked. */
  markup: (name: string) => string;
  read: (markup: string) => unknown;
};

const fromFragment =
  <Model>(read: (element: ReturnType<typeof parseXmlDocument>) => Model) =>
  (markup: string): Model =>
    read(parseXmlDocument(markup));

const paragraphProperties = fromFragment((pPr) => parseParagraphProperties(pPr, null));
const tableProperties = fromFragment(parseTableProperties);
const cellProperties = fromFragment(parseTableCellProperties);

/** Styles reach the model as two records; a fixture that moved neither would prove nothing. */
const styles = (markup: string): unknown => {
  const { styleDefinitions, styles: styleMap } = parseStylesPackage(markup, null);
  return { definitions: styleDefinitions, styles: [...styleMap.entries()] };
};

const numbering = (markup: string): unknown => parseNumbering(markup).definitions;

const stylesPart = (body: string): string => `<w:styles ${WML}>${body}</w:styles>`;

const indent = (name: string): string => `<w:ind w:${name}="720"/>`;
const border = (name: string): string => `<w:${name} w:val="single" w:sz="8" w:color="FF0000"/>`;
const margin = (name: string): string => `<w:${name} w:w="113" w:type="dxa"/>`;

/** The `w:ind` of a paragraph, a style, the document defaults, a conditional region and a list level. */
const INDENT_SITES: readonly ReadSite[] = [
  {
    site: "a paragraph's w:pPr",
    markup: (name) => `<w:pPr ${WML}>${indent(name)}</w:pPr>`,
    read: paragraphProperties,
  },
  {
    site: "a style's w:pPr",
    markup: (name) =>
      stylesPart(
        `<w:style w:type="paragraph" w:styleId="Body"><w:pPr>${indent(name)}</w:pPr></w:style>`,
      ),
    read: styles,
  },
  {
    site: "w:docDefaults/w:pPrDefault",
    markup: (name) =>
      stylesPart(
        `<w:docDefaults><w:pPrDefault><w:pPr>${indent(name)}</w:pPr></w:pPrDefault></w:docDefaults>`,
      ),
    read: styles,
  },
  {
    site: "a table style's w:tblStylePr/w:pPr",
    markup: (name) =>
      stylesPart(
        `<w:style w:type="table" w:styleId="Grid"><w:tblStylePr w:type="firstRow"><w:pPr>${indent(name)}</w:pPr></w:tblStylePr></w:style>`,
      ),
    read: styles,
  },
  {
    site: "a list level's w:pPr",
    markup: (name) =>
      `<w:numbering ${WML}><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:pPr>${indent(name)}</w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`,
    read: numbering,
  },
];

/** `CT_Border` under a table's and a cell's borders, direct and through a style. */
const BORDER_SITES: readonly ReadSite[] = [
  {
    site: "a table's w:tblBorders",
    markup: (name) => `<w:tblPr ${WML}><w:tblBorders>${border(name)}</w:tblBorders></w:tblPr>`,
    read: tableProperties,
  },
  {
    site: "a cell's w:tcBorders",
    markup: (name) => `<w:tcPr ${WML}><w:tcBorders>${border(name)}</w:tcBorders></w:tcPr>`,
    read: cellProperties,
  },
  {
    site: "a table style's w:tblBorders",
    markup: (name) =>
      stylesPart(
        `<w:style w:type="table" w:styleId="Grid"><w:tblPr><w:tblBorders>${border(name)}</w:tblBorders></w:tblPr></w:style>`,
      ),
    read: styles,
  },
  {
    site: "a table style's w:tblStylePr/w:tcPr/w:tcBorders",
    markup: (name) =>
      stylesPart(
        `<w:style w:type="table" w:styleId="Grid"><w:tblStylePr w:type="firstRow"><w:tcPr><w:tcBorders>${border(name)}</w:tcBorders></w:tcPr></w:tblStylePr></w:style>`,
      ),
    read: styles,
  },
];

/** `CT_TblWidth` under a table's and a cell's margins, direct and through a style. */
const MARGIN_SITES: readonly ReadSite[] = [
  {
    site: "a table's w:tblCellMar",
    markup: (name) => `<w:tblPr ${WML}><w:tblCellMar>${margin(name)}</w:tblCellMar></w:tblPr>`,
    read: tableProperties,
  },
  {
    site: "a cell's w:tcMar",
    markup: (name) => `<w:tcPr ${WML}><w:tcMar>${margin(name)}</w:tcMar></w:tcPr>`,
    read: cellProperties,
  },
  {
    site: "a table style's w:tblCellMar",
    markup: (name) =>
      stylesPart(
        `<w:style w:type="table" w:styleId="Grid"><w:tblPr><w:tblCellMar>${margin(name)}</w:tblCellMar></w:tblPr></w:style>`,
      ),
    read: styles,
  },
  {
    site: "a table style's w:tblStylePr/w:tcPr/w:tcMar",
    markup: (name) =>
      stylesPart(
        `<w:style w:type="table" w:styleId="Grid"><w:tblStylePr w:type="firstRow"><w:tcPr><w:tcMar>${margin(name)}</w:tcMar></w:tcPr></w:tblStylePr></w:style>`,
      ),
    read: styles,
  },
];

const READ_SITES = {
  "CT_Border start": BORDER_SITES,
  "CT_Border end": BORDER_SITES,
  "CT_TblWidth start": MARGIN_SITES,
  "CT_TblWidth end": MARGIN_SITES,
  "CT_Ind @start": INDENT_SITES,
  "CT_Ind @end": INDENT_SITES,
  // An indent in character units: folio models neither spelling, so there is no
  // model for the two to agree on. `w:ind|CT_Ind@w:startChars` is a real loss.
  "CT_Ind @startChars": null,
  "CT_Ind @endChars": null,
} as const satisfies Record<StrictName, readonly ReadSite[] | null>;

/** The local name on each side of a rename, without the type the table keys by. */
const localNameOf = (slot: StrictName): string => {
  const local = slot.slice(slot.indexOf(" ") + 1);
  return local.startsWith("@") ? local.slice(1) : local;
};

/**
 * The model with one authored spelling rewritten to another, wherever it is held.
 *
 * A property set folio has not edited is replayed from the bytes it was parsed
 * from, and those bytes carry the name the document really used — correctly, so
 * the capture is the one place the two spellings *should* differ. Rewriting the
 * name rather than deleting the captures says exactly that, and needs no second
 * copy of which slots are captures: a typed field holds a number or a token, so
 * nothing else in the model can match.
 */
const respell = (value: unknown, from: string, to: string): unknown => {
  const name = new RegExp(`w:${from}\\b`, "gu");
  const walk = (current: unknown): unknown => {
    if (typeof current === "string") {
      return current.replaceAll(name, `w:${to}`);
    }
    if (Array.isArray(current)) {
      return current.map(walk);
    }
    if (current instanceof Map) {
      return new Map([...current.entries()].map(([key, held]) => [key, walk(held)]));
    }
    if (typeof current === "object" && current !== null) {
      return Object.fromEntries(
        Object.entries(current).map(([key, held]) => [key, walk(held)] as const),
      );
    }
    return current;
  };
  return walk(value);
};

/** A name no schema declares, so a fixture that ignores its slot reads the same with it. */
const UNREAD_NAME = "folioNoSuchName";

describe("a renamed name reads the same in either spelling", () => {
  for (const slot of STRICT_NAMES) {
    const sites = READ_SITES[slot];
    if (sites === null) {
      continue;
    }
    const strict = localNameOf(slot);
    const transitional = TRANSITIONAL_NAME_BY_STRICT_NAME[slot];

    test(`${slot} at every place folio reads it`, () => {
      for (const { site, markup, read } of sites) {
        const written = read(markup(transitional));
        // A fixture whose slot never reaches the typed model would satisfy the
        // equality below while proving nothing: with the captured bytes
        // rewritten, the two reads would be the same empty model. So the slot
        // has to move something a name the schema never declares does not.
        expect({ site, model: written }).not.toEqual({
          site,
          model: respell(read(markup(UNREAD_NAME)), UNREAD_NAME, transitional),
        });
        expect({ site, model: respell(read(markup(strict)), strict, transitional) }).toEqual({
          site,
          model: written,
        });
      }
    });
  }
});
