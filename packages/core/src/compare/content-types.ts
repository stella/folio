/** How callers expect a block identifier to behave across document revisions. */
export type FolioContentIdStability = "stable" | "positional";

/** One enclosing structural container, ordered outermost to innermost in a block path. */
export type FolioContentContainerPathEntry = {
  kind: string;
  id: string;
};

/** Boolean inline properties represented by content snapshots. */
export type FolioContentInlineBooleanProperty = "bold" | "italic" | "underline" | "strike";

export type FolioContentInlineFormatting = Partial<
  Record<FolioContentInlineBooleanProperty, boolean>
> & {
  fontFamily?: string | null;
  fontSizePt?: number | null;
  color?: string | null;
};

/**
 * An inline-formatting mutation: `false` is an explicit off value, while
 * `null` removes the direct property so its inherited value becomes effective.
 */
export type FolioContentInlineFormattingPatch = Omit<
  FolioContentInlineFormatting,
  FolioContentInlineBooleanProperty
> &
  Partial<Record<FolioContentInlineBooleanProperty, boolean | null>>;

export type FolioContentRun = {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  fontFamily?: string;
  fontSizePt?: number;
  color?: string;
  /** Authored run properties only; paragraph and named-style values stay inherited. */
  directFormatting?: FolioContentInlineFormatting;
};

/**
 * Where a block sits inside its innermost enclosing table. Every index is
 * zero-based: `tableIndex` counts tables in document order across the story,
 * `rowIndex` is the row's index in that table, `cellIndex` is the cell's
 * physical index within the row (a merged cell occupies one slot, so this is
 * not a grid column), and `paragraphIndex` orders the block among the cell's
 * own blocks. Absent on a block that is not inside a table.
 */
export type FolioContentTableLocation = {
  /**
   * Document-order index of the outermost table the block sits in. This equals
   * `tableIndex` unless tables nest.
   */
  outerTableIndex: number;
  tableIndex: number;
  rowIndex: number;
  cellIndex: number;
  /** Grid column occupied by the cell's left edge. */
  gridColumnIndex: number;
  /** Number of grid columns occupied by this physical cell. */
  columnSpan: number;
  /** Number of grid rows occupied by this physical cell. */
  rowSpan: number;
  paragraphIndex: number;
};

/** Direct paragraph alignment understood by the neutral comparison model. */
export type FolioContentParagraphAlignment =
  | "left"
  | "center"
  | "right"
  | "both"
  | "distribute"
  | "mediumKashida"
  | "highKashida"
  | "lowKashida"
  | "thaiDistribute";

/** Line-height interpretation understood by the neutral comparison model. */
export type FolioContentLineSpacingRule = "auto" | "exact" | "atLeast";

/**
 * The complete modeled attribute set of direct paragraph spacing. Optional
 * fields distinguish an absent attribute from an explicit zero or false value.
 */
export type FolioContentParagraphSpacing = {
  spaceBefore?: number;
  spaceAfter?: number;
  lineSpacing?: number;
  lineSpacingRule?: FolioContentLineSpacingRule;
  beforeAutospacing?: boolean;
  afterAutospacing?: boolean;
};

/** A representation-neutral block in one ordered document story. */
export type FolioContentBlock<Kind extends string = string> = {
  id: string;
  kind: Kind;
  text: string;
  idStability?: FolioContentIdStability;
  /** One-based heading depth when the block has outline semantics. */
  headingLevel?: number;
  displayLabel?: string;
  styleId?: string;
  /** Direct paragraph alignment; absent when alignment comes only from a style. */
  directAlignment?: FolioContentParagraphAlignment;
  /** Direct paragraph spacing; absent when every spacing value is inherited. */
  directSpacing?: FolioContentParagraphSpacing;
  /** Zero-based list indent level when the block carries numbering. */
  listLevel?: number;
  previewRuns?: readonly FolioContentRun[];
  table?: FolioContentTableLocation;
  /** Structural ancestry, ordered from the outermost to the innermost container. */
  containerPath?: readonly FolioContentContainerPathEntry[];
};

/** Every block of one story, in document order. */
export type FolioContentSnapshot<
  Block extends FolioContentBlock = FolioContentBlock,
> = {
  blocks: readonly Block[];
};
