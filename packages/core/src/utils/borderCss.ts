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
  type PresetLineDashValue,
  type PresetLineDashVal,
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

const CSS_BORDER_STYLES_DECLARED: ReadonlySet<string> = new Set<string>(CSS_BORDER_STYLE_VALUES);

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
export const cssBorderStyleOf = (value: string | undefined): CssBorderStyle | undefined =>
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
export const CSS_BORDER_STYLES = {
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
export const cssBorderStyle = (style: BorderStyleValue | undefined): CssBorderStyle => {
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
export const cssBorderStyleForDash = (dash: PresetLineDashValue | undefined): CssBorderStyle => {
  if (dash === undefined) {
    return "solid";
  }
  return typeof dash === "string" ? PRESET_DASH_CSS_BORDER_STYLES[dash] : "solid";
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
export const presetDashForCssBorderStyle = (style: string | undefined): PresetLineDashVal =>
  CSS_BORDER_PRESET_DASHES[cssBorderStyleOf(style) ?? "solid"];
