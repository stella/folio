/**
 * The one OOXML border style → CSS `border-style` table.
 *
 * There were three of these and they covered different amounts of the
 * enumeration. `TableExtension` rendered the nine `thinThick`/`thickThin` gap
 * styles as `double` and `dashDotStroked` as `dashed`; the layout bridge and
 * `formatToStyle` fell through to `solid`, so a table cell and the paragraph
 * above it painted the same authored border differently. The three agreed
 * wherever they overlapped, so this table is their union rather than a new
 * decision: a rendering decision per enumeration member is a table, and a table
 * has one owner.
 *
 * Totality is the point of the `satisfies`: `ST_Border` is generated from the
 * committed schema graph, so a schema refresh that adds a member fails the
 * build here rather than quietly painting it as a plain line.
 */

import {
  type BorderStyle,
  type BorderStyleValue,
  type PresetLineDashVal,
  type PresetLineDashValue,
  statesNoBorder,
} from "@stll/docx-core/model";

/** The CSS `border-style` keywords folio paints a border with, in one tuple. */
export const CSS_BORDER_STYLE_VALUES = [
  "none",
  "solid",
  "double",
  "dotted",
  "dashed",
  "groove",
  "ridge",
  "inset",
  "outset",
] as const;

/** The CSS `border-style` keywords folio paints a border with. */
export type CssBorderStyle = (typeof CSS_BORDER_STYLE_VALUES)[number];

const CSS_BORDER_STYLES_DECLARED: ReadonlySet<string> = new Set<string>(
  CSS_BORDER_STYLE_VALUES,
);

/** Whether folio paints a border with this CSS keyword. */
const isCssBorderStyle = (value: string): value is CssBorderStyle =>
  CSS_BORDER_STYLES_DECLARED.has(value);

/**
 * Read the CSS keyword an editor attribute carries into the union folio paints.
 *
 * The attribute is a bare `string` because a host may write any CSS keyword
 * into it. A keyword folio does not paint is not guessed at: the caller gets
 * `undefined` and applies its own default.
 */
export const cssBorderStyleOf = (
  value: string | undefined,
): CssBorderStyle | undefined =>
  value !== undefined && isCssBorderStyle(value) ? value : undefined;

/**
 * Every `ST_Border` member's CSS rendering.
 *
 * The art borders (`apples` … `zigZagStitch`) and the gap-patterned line
 * styles degrade to a plain line, which is how Word itself degrades where the
 * specialised glyphs are unavailable; folio paints the underlying line and
 * round-trips the member. `custom` names art supplied by a relationship, so it
 * degrades the same way.
 */
export const CSS_BORDER_STYLES: {
  readonly nil: "none";
  readonly none: "none";
  readonly single: "solid";
  readonly thick: "solid";
  readonly double: "double";
  readonly dotted: "dotted";
  readonly dashed: "dashed";
  readonly dotDash: "dashed";
  readonly dotDotDash: "dotted";
  readonly triple: "double";
  readonly thinThickSmallGap: "double";
  readonly thickThinSmallGap: "double";
  readonly thinThickThinSmallGap: "double";
  readonly thinThickMediumGap: "double";
  readonly thickThinMediumGap: "double";
  readonly thinThickThinMediumGap: "double";
  readonly thinThickLargeGap: "double";
  readonly thickThinLargeGap: "double";
  readonly thinThickThinLargeGap: "double";
  readonly wave: "solid";
  readonly doubleWave: "double";
  readonly dashSmallGap: "dashed";
  readonly dashDotStroked: "dashed";
  readonly threeDEmboss: "ridge";
  readonly threeDEngrave: "groove";
  readonly outset: "outset";
  readonly inset: "inset";
  readonly apples: "solid";
  readonly archedScallops: "solid";
  readonly babyPacifier: "solid";
  readonly babyRattle: "solid";
  readonly balloons3Colors: "solid";
  readonly balloonsHotAir: "solid";
  readonly basicBlackDashes: "solid";
  readonly basicBlackDots: "solid";
  readonly basicBlackSquares: "solid";
  readonly basicThinLines: "solid";
  readonly basicWhiteDashes: "solid";
  readonly basicWhiteDots: "solid";
  readonly basicWhiteSquares: "solid";
  readonly basicWideInline: "solid";
  readonly basicWideMidline: "solid";
  readonly basicWideOutline: "solid";
  readonly bats: "solid";
  readonly birds: "solid";
  readonly birdsFlight: "solid";
  readonly cabins: "solid";
  readonly cakeSlice: "solid";
  readonly candyCorn: "solid";
  readonly celticKnotwork: "solid";
  readonly certificateBanner: "solid";
  readonly chainLink: "solid";
  readonly champagneBottle: "solid";
  readonly checkedBarBlack: "solid";
  readonly checkedBarColor: "solid";
  readonly checkered: "solid";
  readonly christmasTree: "solid";
  readonly circlesLines: "solid";
  readonly circlesRectangles: "solid";
  readonly classicalWave: "solid";
  readonly clocks: "solid";
  readonly compass: "solid";
  readonly confetti: "solid";
  readonly confettiGrays: "solid";
  readonly confettiOutline: "solid";
  readonly confettiStreamers: "solid";
  readonly confettiWhite: "solid";
  readonly cornerTriangles: "solid";
  readonly couponCutoutDashes: "solid";
  readonly couponCutoutDots: "solid";
  readonly crazyMaze: "solid";
  readonly creaturesButterfly: "solid";
  readonly creaturesFish: "solid";
  readonly creaturesInsects: "solid";
  readonly creaturesLadyBug: "solid";
  readonly crossStitch: "solid";
  readonly cup: "solid";
  readonly decoArch: "solid";
  readonly decoArchColor: "solid";
  readonly decoBlocks: "solid";
  readonly diamondsGray: "solid";
  readonly doubleD: "solid";
  readonly doubleDiamonds: "solid";
  readonly earth1: "solid";
  readonly earth2: "solid";
  readonly earth3: "solid";
  readonly eclipsingSquares1: "solid";
  readonly eclipsingSquares2: "solid";
  readonly eggsBlack: "solid";
  readonly fans: "solid";
  readonly film: "solid";
  readonly firecrackers: "solid";
  readonly flowersBlockPrint: "solid";
  readonly flowersDaisies: "solid";
  readonly flowersModern1: "solid";
  readonly flowersModern2: "solid";
  readonly flowersPansy: "solid";
  readonly flowersRedRose: "solid";
  readonly flowersRoses: "solid";
  readonly flowersTeacup: "solid";
  readonly flowersTiny: "solid";
  readonly gems: "solid";
  readonly gingerbreadMan: "solid";
  readonly gradient: "solid";
  readonly handmade1: "solid";
  readonly handmade2: "solid";
  readonly heartBalloon: "solid";
  readonly heartGray: "solid";
  readonly hearts: "solid";
  readonly heebieJeebies: "solid";
  readonly holly: "solid";
  readonly houseFunky: "solid";
  readonly hypnotic: "solid";
  readonly iceCreamCones: "solid";
  readonly lightBulb: "solid";
  readonly lightning1: "solid";
  readonly lightning2: "solid";
  readonly mapPins: "solid";
  readonly mapleLeaf: "solid";
  readonly mapleMuffins: "solid";
  readonly marquee: "solid";
  readonly marqueeToothed: "solid";
  readonly moons: "solid";
  readonly mosaic: "solid";
  readonly musicNotes: "solid";
  readonly northwest: "solid";
  readonly ovals: "solid";
  readonly packages: "solid";
  readonly palmsBlack: "solid";
  readonly palmsColor: "solid";
  readonly paperClips: "solid";
  readonly papyrus: "solid";
  readonly partyFavor: "solid";
  readonly partyGlass: "solid";
  readonly pencils: "solid";
  readonly people: "solid";
  readonly peopleWaving: "solid";
  readonly peopleHats: "solid";
  readonly poinsettias: "solid";
  readonly postageStamp: "solid";
  readonly pumpkin1: "solid";
  readonly pushPinNote2: "solid";
  readonly pushPinNote1: "solid";
  readonly pyramids: "solid";
  readonly pyramidsAbove: "solid";
  readonly quadrants: "solid";
  readonly rings: "solid";
  readonly safari: "solid";
  readonly sawtooth: "solid";
  readonly sawtoothGray: "solid";
  readonly scaredCat: "solid";
  readonly seattle: "solid";
  readonly shadowedSquares: "solid";
  readonly sharksTeeth: "solid";
  readonly shorebirdTracks: "solid";
  readonly skyrocket: "solid";
  readonly snowflakeFancy: "solid";
  readonly snowflakes: "solid";
  readonly sombrero: "solid";
  readonly southwest: "solid";
  readonly stars: "solid";
  readonly starsTop: "solid";
  readonly stars3d: "solid";
  readonly starsBlack: "solid";
  readonly starsShadowed: "solid";
  readonly sun: "solid";
  readonly swirligig: "solid";
  readonly tornPaper: "solid";
  readonly tornPaperBlack: "solid";
  readonly trees: "solid";
  readonly triangleParty: "solid";
  readonly triangles: "solid";
  readonly triangle1: "solid";
  readonly triangle2: "solid";
  readonly triangleCircle1: "solid";
  readonly triangleCircle2: "solid";
  readonly shapes1: "solid";
  readonly shapes2: "solid";
  readonly twistedLines1: "solid";
  readonly twistedLines2: "solid";
  readonly vine: "solid";
  readonly waveline: "solid";
  readonly weavingAngles: "solid";
  readonly weavingBraid: "solid";
  readonly weavingRibbon: "solid";
  readonly weavingStrips: "solid";
  readonly whiteFlowers: "solid";
  readonly woodwork: "solid";
  readonly xIllusions: "solid";
  readonly zanyTriangles: "solid";
  readonly zigZag: "solid";
  readonly zigZagStitch: "solid";
  readonly custom: "solid";
} = {
  nil: "none",
  none: "none",
  single: "solid",
  thick: "solid",
  double: "double",
  dotted: "dotted",
  dashed: "dashed",
  dotDash: "dashed",
  dotDotDash: "dotted",
  triple: "double",
  thinThickSmallGap: "double",
  thickThinSmallGap: "double",
  thinThickThinSmallGap: "double",
  thinThickMediumGap: "double",
  thickThinMediumGap: "double",
  thinThickThinMediumGap: "double",
  thinThickLargeGap: "double",
  thickThinLargeGap: "double",
  thinThickThinLargeGap: "double",
  wave: "solid",
  doubleWave: "double",
  dashSmallGap: "dashed",
  dashDotStroked: "dashed",
  threeDEmboss: "ridge",
  threeDEngrave: "groove",
  outset: "outset",
  inset: "inset",
  apples: "solid",
  archedScallops: "solid",
  babyPacifier: "solid",
  babyRattle: "solid",
  balloons3Colors: "solid",
  balloonsHotAir: "solid",
  basicBlackDashes: "solid",
  basicBlackDots: "solid",
  basicBlackSquares: "solid",
  basicThinLines: "solid",
  basicWhiteDashes: "solid",
  basicWhiteDots: "solid",
  basicWhiteSquares: "solid",
  basicWideInline: "solid",
  basicWideMidline: "solid",
  basicWideOutline: "solid",
  bats: "solid",
  birds: "solid",
  birdsFlight: "solid",
  cabins: "solid",
  cakeSlice: "solid",
  candyCorn: "solid",
  celticKnotwork: "solid",
  certificateBanner: "solid",
  chainLink: "solid",
  champagneBottle: "solid",
  checkedBarBlack: "solid",
  checkedBarColor: "solid",
  checkered: "solid",
  christmasTree: "solid",
  circlesLines: "solid",
  circlesRectangles: "solid",
  classicalWave: "solid",
  clocks: "solid",
  compass: "solid",
  confetti: "solid",
  confettiGrays: "solid",
  confettiOutline: "solid",
  confettiStreamers: "solid",
  confettiWhite: "solid",
  cornerTriangles: "solid",
  couponCutoutDashes: "solid",
  couponCutoutDots: "solid",
  crazyMaze: "solid",
  creaturesButterfly: "solid",
  creaturesFish: "solid",
  creaturesInsects: "solid",
  creaturesLadyBug: "solid",
  crossStitch: "solid",
  cup: "solid",
  decoArch: "solid",
  decoArchColor: "solid",
  decoBlocks: "solid",
  diamondsGray: "solid",
  doubleD: "solid",
  doubleDiamonds: "solid",
  earth1: "solid",
  earth2: "solid",
  earth3: "solid",
  eclipsingSquares1: "solid",
  eclipsingSquares2: "solid",
  eggsBlack: "solid",
  fans: "solid",
  film: "solid",
  firecrackers: "solid",
  flowersBlockPrint: "solid",
  flowersDaisies: "solid",
  flowersModern1: "solid",
  flowersModern2: "solid",
  flowersPansy: "solid",
  flowersRedRose: "solid",
  flowersRoses: "solid",
  flowersTeacup: "solid",
  flowersTiny: "solid",
  gems: "solid",
  gingerbreadMan: "solid",
  gradient: "solid",
  handmade1: "solid",
  handmade2: "solid",
  heartBalloon: "solid",
  heartGray: "solid",
  hearts: "solid",
  heebieJeebies: "solid",
  holly: "solid",
  houseFunky: "solid",
  hypnotic: "solid",
  iceCreamCones: "solid",
  lightBulb: "solid",
  lightning1: "solid",
  lightning2: "solid",
  mapPins: "solid",
  mapleLeaf: "solid",
  mapleMuffins: "solid",
  marquee: "solid",
  marqueeToothed: "solid",
  moons: "solid",
  mosaic: "solid",
  musicNotes: "solid",
  northwest: "solid",
  ovals: "solid",
  packages: "solid",
  palmsBlack: "solid",
  palmsColor: "solid",
  paperClips: "solid",
  papyrus: "solid",
  partyFavor: "solid",
  partyGlass: "solid",
  pencils: "solid",
  people: "solid",
  peopleWaving: "solid",
  peopleHats: "solid",
  poinsettias: "solid",
  postageStamp: "solid",
  pumpkin1: "solid",
  pushPinNote2: "solid",
  pushPinNote1: "solid",
  pyramids: "solid",
  pyramidsAbove: "solid",
  quadrants: "solid",
  rings: "solid",
  safari: "solid",
  sawtooth: "solid",
  sawtoothGray: "solid",
  scaredCat: "solid",
  seattle: "solid",
  shadowedSquares: "solid",
  sharksTeeth: "solid",
  shorebirdTracks: "solid",
  skyrocket: "solid",
  snowflakeFancy: "solid",
  snowflakes: "solid",
  sombrero: "solid",
  southwest: "solid",
  stars: "solid",
  starsTop: "solid",
  stars3d: "solid",
  starsBlack: "solid",
  starsShadowed: "solid",
  sun: "solid",
  swirligig: "solid",
  tornPaper: "solid",
  tornPaperBlack: "solid",
  trees: "solid",
  triangleParty: "solid",
  triangles: "solid",
  triangle1: "solid",
  triangle2: "solid",
  triangleCircle1: "solid",
  triangleCircle2: "solid",
  shapes1: "solid",
  shapes2: "solid",
  twistedLines1: "solid",
  twistedLines2: "solid",
  vine: "solid",
  waveline: "solid",
  weavingAngles: "solid",
  weavingBraid: "solid",
  weavingRibbon: "solid",
  weavingStrips: "solid",
  whiteFlowers: "solid",
  woodwork: "solid",
  xIllusions: "solid",
  zanyTriangles: "solid",
  zigZag: "solid",
  zigZagStitch: "solid",
  custom: "solid",
} as const satisfies Record<BorderStyle, CssBorderStyle>;

/** The CSS keyword for an edge that paints nothing. */
const CSS_NO_BORDER = "none" as const satisfies CssBorderStyle;

/**
 * Whether a laid-out border paints a line.
 *
 * Only an explicit `none` does not: a border with a width but no style paints
 * the CSS initial `solid`, which is what the measurers and the painter have
 * always assumed. The argument is a CSS keyword, never an `ST_Border` member —
 * `nil` cannot reach here, which is what typing `BorderStyle.style` bought.
 */
export const paintsCssBorder = (style: CssBorderStyle | undefined): boolean =>
  style !== CSS_NO_BORDER;

/**
 * The CSS `border-style` for a parsed `w:val`.
 *
 * A token the schema does not declare paints as a plain line: the member is
 * kept for the round trip, but nothing can be inferred about how it draws.
 */
export const cssBorderStyle = (
  style: BorderStyleValue | undefined,
): CssBorderStyle => {
  if (style === undefined) {
    return "solid";
  }
  if (statesNoBorder(style)) {
    return "none";
  }
  return typeof style === "string" ? CSS_BORDER_STYLES[style] : "solid";
};

/**
 * Every `ST_PresetLineDashVal` member's CSS `border-style` rendering.
 *
 * A shape or text-box outline reaches the DOM painter and the editor node as a
 * CSS border, so the dash has to be translated rather than forwarded: `sysDash`
 * in a `border` shorthand invalidates the whole declaration and the outline
 * disappears. CSS has three periodic keywords, so the dash-length and system
 * variants collapse; the model keeps the authored member for the round trip.
 *
 * Not the inverse of {@link CSS_BORDER_PRESET_DASHES}: both directions are
 * many-to-one, so neither can be derived from the other.
 */
const PRESET_DASH_CSS_BORDER_STYLES = {
  solid: "solid",
  dot: "dotted",
  dash: "dashed",
  lgDash: "dashed",
  dashDot: "dashed",
  lgDashDot: "dashed",
  lgDashDotDot: "dashed",
  sysDash: "dashed",
  sysDot: "dotted",
  sysDashDot: "dashed",
  sysDashDotDot: "dashed",
} as const satisfies Record<PresetLineDashVal, CssBorderStyle>;

/**
 * The CSS `border-style` for a parsed `a:prstDash@val`.
 *
 * A token the schema does not declare paints as a plain line, the way an
 * undeclared `ST_Border` member does: the member is kept for the round trip,
 * but nothing can be inferred about how it draws.
 */
export const cssBorderStyleForDash = (
  dash: PresetLineDashValue | undefined,
): CssBorderStyle => {
  if (dash === undefined) {
    return "solid";
  }
  return typeof dash === "string"
    ? PRESET_DASH_CSS_BORDER_STYLES[dash]
    : "solid";
};

/**
 * Every CSS `border-style` folio paints, as the preset dash it serializes to.
 *
 * An image border is authored in CSS on the editor node and written back as an
 * `a:ln` outline, so the translation has to be total: a keyword with no entry
 * was serialized as `solid` by a partial table that could not say so.
 */
const CSS_BORDER_PRESET_DASHES = {
  none: "solid",
  solid: "solid",
  double: "solid",
  dotted: "dot",
  dashed: "dash",
  groove: "solid",
  ridge: "solid",
  inset: "solid",
  outset: "solid",
} as const satisfies Record<CssBorderStyle, PresetLineDashVal>;

/** The preset dash an authored CSS border style serializes to. */
export const presetDashForCssBorderStyle = (
  style: string | undefined,
): PresetLineDashVal =>
  CSS_BORDER_PRESET_DASHES[cssBorderStyleOf(style) ?? "solid"];
