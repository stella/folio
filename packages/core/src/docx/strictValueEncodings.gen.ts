/**
 * GENERATED FILE — do not edit.
 *
 * How a Strict-spelled measure or percentage is written under a Transitional
 * root, derived from the Transitional schema graph by
 * `scripts/generate-strict-value-encodings.ts`. Regenerate with:
 *
 *   bun run generate:strict-value-encodings
 */

import type { MeasureUnit } from "./universalMeasure";

/** The fraction of a percent a Transitional attribute's number counts. */
export type PercentUnit = "fiftiethPercent" | "thousandthPercent" | "wholePercent";

/** The number a slot's dual-spelling type accepts in place of a Strict string. */
export type SlotEncoding = {
  readonly measure?: MeasureUnit;
  readonly percent?: PercentUnit;
};

const MEASURE_UNITS: Readonly<Record<string, MeasureUnit>> = {
  emu: "emu",
  halfPoints: "halfPoints",
  hundredthPoints: "hundredthPoints",
  twips: "twips",
};

const PERCENT_UNITS: Readonly<Record<string, PercentUnit>> = {
  fiftiethPercent: "fiftiethPercent",
  thousandthPercent: "thousandthPercent",
  wholePercent: "wholePercent",
};

/**
 * One slot per line: the slot, a tab, its measure unit, a tab, its percent unit.
 *
 * Text rather than object literals because every package that depends on
 * `@stll/folio-core` pays this file's inference cost, and a few hundred
 * literals breach the repository's compiler-workload budget on their own.
 * The slot is `"<namespace URI> <element local name>"` for element text and
 * `"<namespace URI> <element local name> @<attribute local name>"` for an
 * attribute; an empty column means the type has no spelling of that kind.
 */
const SLOT_TABLE = `http://schemas.openxmlformats.org/drawingml/2006/main ahPolar @maxR	emu	
http://schemas.openxmlformats.org/drawingml/2006/main ahPolar @minR	emu	
http://schemas.openxmlformats.org/drawingml/2006/main ahXY @maxX	emu	
http://schemas.openxmlformats.org/drawingml/2006/main ahXY @maxY	emu	
http://schemas.openxmlformats.org/drawingml/2006/main ahXY @minX	emu	
http://schemas.openxmlformats.org/drawingml/2006/main ahXY @minY	emu	
http://schemas.openxmlformats.org/drawingml/2006/main alpha @val		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main alphaBiLevel @thresh		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main alphaMod @val		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main alphaModFix @amt		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main alphaOff @val		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main alphaOutset @rad	emu	
http://schemas.openxmlformats.org/drawingml/2006/main alphaRepl @a		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main anchor @x	emu	
http://schemas.openxmlformats.org/drawingml/2006/main anchor @y	emu	
http://schemas.openxmlformats.org/drawingml/2006/main anchor @z	emu	
http://schemas.openxmlformats.org/drawingml/2006/main arcTo @hR	emu	
http://schemas.openxmlformats.org/drawingml/2006/main arcTo @wR	emu	
http://schemas.openxmlformats.org/drawingml/2006/main biLevel @thresh		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main blue @val		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main blueMod @val		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main blueOff @val		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main bodyPr @bIns	emu	
http://schemas.openxmlformats.org/drawingml/2006/main bodyPr @lIns	emu	
http://schemas.openxmlformats.org/drawingml/2006/main bodyPr @rIns	emu	
http://schemas.openxmlformats.org/drawingml/2006/main bodyPr @tIns	emu	
http://schemas.openxmlformats.org/drawingml/2006/main camera @zoom		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main chOff @x	emu	
http://schemas.openxmlformats.org/drawingml/2006/main chOff @y	emu	
http://schemas.openxmlformats.org/drawingml/2006/main defPPr @defTabSz	emu	
http://schemas.openxmlformats.org/drawingml/2006/main defRPr @baseline		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main defRPr @spc	hundredthPoints	
http://schemas.openxmlformats.org/drawingml/2006/main ds @d		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main ds @sp		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main endParaRPr @baseline		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main endParaRPr @spc	hundredthPoints	
http://schemas.openxmlformats.org/drawingml/2006/main fillRect @b		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main fillRect @l		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main fillRect @r		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main fillRect @t		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main fillToRect @b		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main fillToRect @l		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main fillToRect @r		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main fillToRect @t		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main flatTx @z	emu	
http://schemas.openxmlformats.org/drawingml/2006/main green @val		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main greenMod @val		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main greenOff @val		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main gridCol @w	emu	
http://schemas.openxmlformats.org/drawingml/2006/main gs @pos		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main hsl @lum		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main hsl @sat		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main hslClr @lum		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main hslClr @sat		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main hueMod @val		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main lum @bright		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main lum @contrast		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main lum @val		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main lumMod @val		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main lumOff @val		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main lvl1pPr @defTabSz	emu	
http://schemas.openxmlformats.org/drawingml/2006/main lvl2pPr @defTabSz	emu	
http://schemas.openxmlformats.org/drawingml/2006/main lvl3pPr @defTabSz	emu	
http://schemas.openxmlformats.org/drawingml/2006/main lvl4pPr @defTabSz	emu	
http://schemas.openxmlformats.org/drawingml/2006/main lvl5pPr @defTabSz	emu	
http://schemas.openxmlformats.org/drawingml/2006/main lvl6pPr @defTabSz	emu	
http://schemas.openxmlformats.org/drawingml/2006/main lvl7pPr @defTabSz	emu	
http://schemas.openxmlformats.org/drawingml/2006/main lvl8pPr @defTabSz	emu	
http://schemas.openxmlformats.org/drawingml/2006/main lvl9pPr @defTabSz	emu	
http://schemas.openxmlformats.org/drawingml/2006/main miter @lim		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main norm @dx	emu	
http://schemas.openxmlformats.org/drawingml/2006/main norm @dy	emu	
http://schemas.openxmlformats.org/drawingml/2006/main norm @dz	emu	
http://schemas.openxmlformats.org/drawingml/2006/main normAutofit @fontScale		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main normAutofit @lnSpcReduction		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main off @x	emu	
http://schemas.openxmlformats.org/drawingml/2006/main off @y	emu	
http://schemas.openxmlformats.org/drawingml/2006/main outerShdw @sx		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main outerShdw @sy		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main pPr @defTabSz	emu	
http://schemas.openxmlformats.org/drawingml/2006/main pos @x	emu	
http://schemas.openxmlformats.org/drawingml/2006/main pos @y	emu	
http://schemas.openxmlformats.org/drawingml/2006/main pt @x	emu	
http://schemas.openxmlformats.org/drawingml/2006/main pt @y	emu	
http://schemas.openxmlformats.org/drawingml/2006/main rPr @baseline		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main rPr @spc	hundredthPoints	
http://schemas.openxmlformats.org/drawingml/2006/main rect @b	emu	
http://schemas.openxmlformats.org/drawingml/2006/main rect @l	emu	
http://schemas.openxmlformats.org/drawingml/2006/main rect @r	emu	
http://schemas.openxmlformats.org/drawingml/2006/main rect @t	emu	
http://schemas.openxmlformats.org/drawingml/2006/main red @val		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main redMod @val		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main redOff @val		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main reflection @endA		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main reflection @endPos		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main reflection @stA		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main reflection @stPos		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main reflection @sx		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main reflection @sy		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main relOff @tx		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main relOff @ty		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main sat @val		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main satMod @val		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main satOff @val		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main scrgbClr @b		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main scrgbClr @g		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main scrgbClr @r		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main shade @val		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main sp3d @z	emu	
http://schemas.openxmlformats.org/drawingml/2006/main spcPct @val		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main srcRect @b		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main srcRect @l		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main srcRect @r		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main srcRect @t		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main tab @pos	emu	
http://schemas.openxmlformats.org/drawingml/2006/main tcPr @marB	emu	
http://schemas.openxmlformats.org/drawingml/2006/main tcPr @marL	emu	
http://schemas.openxmlformats.org/drawingml/2006/main tcPr @marR	emu	
http://schemas.openxmlformats.org/drawingml/2006/main tcPr @marT	emu	
http://schemas.openxmlformats.org/drawingml/2006/main tile @sx		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main tile @sy		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main tile @tx	emu	
http://schemas.openxmlformats.org/drawingml/2006/main tile @ty	emu	
http://schemas.openxmlformats.org/drawingml/2006/main tileRect @b		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main tileRect @l		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main tileRect @r		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main tileRect @t		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main tint @amt		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main tint @val		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main tr @h	emu	
http://schemas.openxmlformats.org/drawingml/2006/main up @dx	emu	
http://schemas.openxmlformats.org/drawingml/2006/main up @dy	emu	
http://schemas.openxmlformats.org/drawingml/2006/main up @dz	emu	
http://schemas.openxmlformats.org/drawingml/2006/main xfrm @sx		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main xfrm @sy		thousandthPercent
http://schemas.openxmlformats.org/drawingml/2006/main xfrm @tx	emu	
http://schemas.openxmlformats.org/drawingml/2006/main xfrm @ty	emu	
http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing bodyPr @bIns	emu	
http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing bodyPr @lIns	emu	
http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing bodyPr @rIns	emu	
http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing bodyPr @tIns	emu	
http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing effectExtent @b	emu	
http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing effectExtent @l	emu	
http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing effectExtent @r	emu	
http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing effectExtent @t	emu	
http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing lineTo @x	emu	
http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing lineTo @y	emu	
http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing simplePos @x	emu	
http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing simplePos @y	emu	
http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing start @x	emu	
http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing start @y	emu	
http://schemas.openxmlformats.org/officeDocument/2006/math interSp @val	twips	
http://schemas.openxmlformats.org/officeDocument/2006/math intraSp @val	twips	
http://schemas.openxmlformats.org/officeDocument/2006/math lMargin @val	twips	
http://schemas.openxmlformats.org/officeDocument/2006/math postSp @val	twips	
http://schemas.openxmlformats.org/officeDocument/2006/math preSp @val	twips	
http://schemas.openxmlformats.org/officeDocument/2006/math rMargin @val	twips	
http://schemas.openxmlformats.org/officeDocument/2006/math wrapIndent @val	twips	
http://schemas.openxmlformats.org/wordprocessingml/2006/main bottom @w	twips	fiftiethPercent
http://schemas.openxmlformats.org/wordprocessingml/2006/main col @space	twips	
http://schemas.openxmlformats.org/wordprocessingml/2006/main col @w	twips	
http://schemas.openxmlformats.org/wordprocessingml/2006/main cols @space	twips	
http://schemas.openxmlformats.org/wordprocessingml/2006/main end @w	twips	fiftiethPercent
http://schemas.openxmlformats.org/wordprocessingml/2006/main fitText @val	twips	
http://schemas.openxmlformats.org/wordprocessingml/2006/main framePr @h	twips	
http://schemas.openxmlformats.org/wordprocessingml/2006/main framePr @hSpace	twips	
http://schemas.openxmlformats.org/wordprocessingml/2006/main framePr @vSpace	twips	
http://schemas.openxmlformats.org/wordprocessingml/2006/main framePr @w	twips	
http://schemas.openxmlformats.org/wordprocessingml/2006/main framePr @x	twips	
http://schemas.openxmlformats.org/wordprocessingml/2006/main framePr @y	twips	
http://schemas.openxmlformats.org/wordprocessingml/2006/main gridCol @w	twips	
http://schemas.openxmlformats.org/wordprocessingml/2006/main hps @val	halfPoints	
http://schemas.openxmlformats.org/wordprocessingml/2006/main hpsBaseText @val	halfPoints	
http://schemas.openxmlformats.org/wordprocessingml/2006/main hpsRaise @val	halfPoints	
http://schemas.openxmlformats.org/wordprocessingml/2006/main ind @end	twips	
http://schemas.openxmlformats.org/wordprocessingml/2006/main ind @firstLine	twips	
http://schemas.openxmlformats.org/wordprocessingml/2006/main ind @hanging	twips	
http://schemas.openxmlformats.org/wordprocessingml/2006/main ind @left	twips	
http://schemas.openxmlformats.org/wordprocessingml/2006/main ind @right	twips	
http://schemas.openxmlformats.org/wordprocessingml/2006/main ind @start	twips	
http://schemas.openxmlformats.org/wordprocessingml/2006/main kern @val	halfPoints	
http://schemas.openxmlformats.org/wordprocessingml/2006/main left @w	twips	fiftiethPercent
http://schemas.openxmlformats.org/wordprocessingml/2006/main lnNumType @distance	twips	
http://schemas.openxmlformats.org/wordprocessingml/2006/main object @dxaOrig	twips	
http://schemas.openxmlformats.org/wordprocessingml/2006/main object @dyaOrig	twips	
http://schemas.openxmlformats.org/wordprocessingml/2006/main pgMar @bottom	twips	
http://schemas.openxmlformats.org/wordprocessingml/2006/main pgMar @footer	twips	
http://schemas.openxmlformats.org/wordprocessingml/2006/main pgMar @gutter	twips	
http://schemas.openxmlformats.org/wordprocessingml/2006/main pgMar @header	twips	
http://schemas.openxmlformats.org/wordprocessingml/2006/main pgMar @left	twips	
http://schemas.openxmlformats.org/wordprocessingml/2006/main pgMar @right	twips	
http://schemas.openxmlformats.org/wordprocessingml/2006/main pgMar @top	twips	
http://schemas.openxmlformats.org/wordprocessingml/2006/main pgSz @h	twips	
http://schemas.openxmlformats.org/wordprocessingml/2006/main pgSz @w	twips	
http://schemas.openxmlformats.org/wordprocessingml/2006/main position @val	halfPoints	
http://schemas.openxmlformats.org/wordprocessingml/2006/main right @w	twips	fiftiethPercent
http://schemas.openxmlformats.org/wordprocessingml/2006/main size @val	halfPoints	
http://schemas.openxmlformats.org/wordprocessingml/2006/main spacing @after	twips	
http://schemas.openxmlformats.org/wordprocessingml/2006/main spacing @before	twips	
http://schemas.openxmlformats.org/wordprocessingml/2006/main spacing @line	twips	
http://schemas.openxmlformats.org/wordprocessingml/2006/main spacing @val	twips	
http://schemas.openxmlformats.org/wordprocessingml/2006/main start @w	twips	fiftiethPercent
http://schemas.openxmlformats.org/wordprocessingml/2006/main sz @val	halfPoints	
http://schemas.openxmlformats.org/wordprocessingml/2006/main szCs @val	halfPoints	
http://schemas.openxmlformats.org/wordprocessingml/2006/main tab @pos	twips	
http://schemas.openxmlformats.org/wordprocessingml/2006/main tblCellSpacing @w	twips	fiftiethPercent
http://schemas.openxmlformats.org/wordprocessingml/2006/main tblInd @w	twips	fiftiethPercent
http://schemas.openxmlformats.org/wordprocessingml/2006/main tblW @w	twips	fiftiethPercent
http://schemas.openxmlformats.org/wordprocessingml/2006/main tblpPr @bottomFromText	twips	
http://schemas.openxmlformats.org/wordprocessingml/2006/main tblpPr @leftFromText	twips	
http://schemas.openxmlformats.org/wordprocessingml/2006/main tblpPr @rightFromText	twips	
http://schemas.openxmlformats.org/wordprocessingml/2006/main tblpPr @tblpX	twips	
http://schemas.openxmlformats.org/wordprocessingml/2006/main tblpPr @tblpY	twips	
http://schemas.openxmlformats.org/wordprocessingml/2006/main tblpPr @topFromText	twips	
http://schemas.openxmlformats.org/wordprocessingml/2006/main tcW @w	twips	fiftiethPercent
http://schemas.openxmlformats.org/wordprocessingml/2006/main top @w	twips	fiftiethPercent
http://schemas.openxmlformats.org/wordprocessingml/2006/main trHeight @val	twips	
http://schemas.openxmlformats.org/wordprocessingml/2006/main w @val		wholePercent
http://schemas.openxmlformats.org/wordprocessingml/2006/main wAfter @w	twips	fiftiethPercent
http://schemas.openxmlformats.org/wordprocessingml/2006/main wBefore @w	twips	fiftiethPercent`;

const readSlotTable = (): ReadonlyMap<string, SlotEncoding> => {
  const slots = new Map<string, SlotEncoding>();
  for (const line of SLOT_TABLE.split("\n")) {
    const [slot, measure, percent] = line.split("\t");
    if (slot === undefined) {
      continue;
    }
    const measureUnit = measure === undefined ? undefined : MEASURE_UNITS[measure];
    const percentUnit = percent === undefined ? undefined : PERCENT_UNITS[percent];
    slots.set(slot, {
      ...(measureUnit === undefined ? {} : { measure: measureUnit }),
      ...(percentUnit === undefined ? {} : { percent: percentUnit }),
    });
  }
  return slots;
};

/** Slots whose Transitional type spells one value two ways. */
export const TRANSITIONAL_SLOT_ENCODINGS: ReadonlyMap<string, SlotEncoding> = readSlotTable();

const NAMESPACE_PAIRS: readonly (readonly [strict: string, transitional: string])[] = [
  ["http://purl.oclc.org/ooxml/drawingml/chart", "http://schemas.openxmlformats.org/drawingml/2006/chart"],
  ["http://purl.oclc.org/ooxml/drawingml/chartDrawing", "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing"],
  ["http://purl.oclc.org/ooxml/drawingml/diagram", "http://schemas.openxmlformats.org/drawingml/2006/diagram"],
  ["http://purl.oclc.org/ooxml/drawingml/lockedCanvas", "http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas"],
  ["http://purl.oclc.org/ooxml/drawingml/main", "http://schemas.openxmlformats.org/drawingml/2006/main"],
  ["http://purl.oclc.org/ooxml/drawingml/picture", "http://schemas.openxmlformats.org/drawingml/2006/picture"],
  ["http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"],
  ["http://purl.oclc.org/ooxml/officeDocument/math", "http://schemas.openxmlformats.org/officeDocument/2006/math"],
  ["http://purl.oclc.org/ooxml/officeDocument/relationships", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"],
  ["http://purl.oclc.org/ooxml/officeDocument/sharedTypes", "http://schemas.openxmlformats.org/officeDocument/2006/sharedTypes"],
  ["http://purl.oclc.org/ooxml/schemaLibrary/main", "http://schemas.openxmlformats.org/schemaLibrary/2006/main"],
  ["http://purl.oclc.org/ooxml/wordprocessingml/main", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"],
];

/** Every Strict namespace URI a WordprocessingML part can carry, and its Transitional pair. */
export const TRANSITIONAL_NAMESPACE_BY_STRICT_URI: ReadonlyMap<string, string> = new Map(
  NAMESPACE_PAIRS,
);
