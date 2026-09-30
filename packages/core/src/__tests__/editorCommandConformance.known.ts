/**
 * Known gaps in the editor-command conformance matrix.
 *
 * Each entry names the cases it covers and the violation kinds it expects.
 * A covered case that shows an expected violation passes; a violation no
 * entry covers fails. An entry that covered cases in a run but saw none of
 * its violations is stale and fails too, so a fix removes its entry in the
 * same change — entries tied to an issue flip that way when the fix lands.
 */

import type { ViolationKind } from "./editorCommandConformance";
import type { EditorMode, SelectionPlacement } from "./editorHarness";

export type ConformanceCaseKey = {
  shape: string;
  operation: string;
  placement: SelectionPlacement;
};

export type KnownConformanceGap = {
  /** The tracking issue, when there is one. */
  issue?: number;
  reason: string;
  /** Operations covered; every operation when omitted. */
  operations?: readonly string[];
  /** Shapes covered; every shape when omitted. */
  shapes?: readonly string[];
  /** Placements covered; every placement when omitted. */
  placements?: readonly SelectionPlacement[];
  /** Shape/operation pairs that have complete resolution coverage. */
  excludedCases?: readonly Pick<ConformanceCaseKey, "shape" | "operation">[];
  modes?: readonly EditorMode[];
  kinds: readonly ViolationKind[];
  /**
   * `full` for a gap only the full tier's placements reach; the default tier
   * then does not hold it to the staleness check.
   */
  tier?: "full";
};

const LIST_TOGGLES = ["command:toggleBulletList", "command:toggleNumberedList"] as const;
const LIST_MARKERS = ["type:bullet-marker", "type:star-marker", "type:number-marker"] as const;
const LIST_NUMBERING = [
  "command:restartNumbering",
  "command:continueNumbering",
  "command:setNumberingValue(3)",
] as const;
/** Operations that type, insert or delete at the selection, replacing a range when there is one. */
const REPLACING_OPERATIONS = [
  "type:text",
  "type:text(mid)",
  ...LIST_MARKERS,
  "key:Enter",
  "key:Shift-Enter",
  "key:Mod-Enter",
  "key:Tab",
  "key:Backspace",
  "key:Delete",
  "key:Mod-Backspace",
  "key:Mod-Delete",
  "key:Shift-Backspace",
  "command:insertHyperlink",
  "command:insertFootnote",
  "command:insertEndnote",
  "paste:paragraphs",
  "paste:copied-blocks",
] as const;

const RANGE_PLACEMENTS: readonly SelectionPlacement[] = [
  "word",
  "paragraph",
  "cross-paragraph",
  "document",
];

export const KNOWN_CONFORMANCE_GAPS: readonly KnownConformanceGap[] = [
  // ---------------------------------------------------------------- lists --
  {
    reason:
      "List operations in suggesting mode: rejecting a tracked toggle or autoformat writes resolved spacing and indentation back as direct formatting; level changes and list removal apply untracked",
    operations: [
      ...LIST_TOGGLES,
      ...LIST_MARKERS,
      ...LIST_NUMBERING,
      "command:increaseListLevel",
      "command:decreaseListLevel",
      "command:removeList",
      "key:Tab",
      "key:Shift-Tab",
    ],
    modes: ["suggesting"],
    kinds: ["reject-mismatch"],
  },
  {
    reason:
      "Modified Backspace at the start of a list item joins the next item's text directly, so rejecting its suggestion does not restore the original text",
    operations: ["key:Mod-Backspace", "key:Shift-Backspace"],
    shapes: ["single-decimal-list", "single-bullet-list", "outline-level-numbered"],
    placements: ["caret-start"],
    modes: ["suggesting"],
    kinds: ["reject-mismatch"],
  },

  // ---------------------------------------------------- suggesting mode --
  {
    reason:
      "Suggesting mode applies character and paragraph formatting, paragraph styles and hyperlinks directly instead of recording w:rPrChange / w:pPrChange",
    operations: [
      "command:toggleBold",
      "command:toggleItalic",
      "command:toggleUnderline",
      "command:toggleStrike",
      "command:setTextColor",
      "command:setFontSize",
      "command:setFontFamily",
      "command:setUnderlineStyle",
      "command:setHyperlink",
      "key:Mod-b",
      "key:Mod-i",
      "key:Mod-u",
      "command:setAlignment(center)",
      "command:alignLeft",
      "command:alignCenter",
      "command:alignRight",
      "command:alignJustify",
      "command:setLineSpacing(360)",
      "command:singleSpacing",
      "command:oneAndHalfSpacing",
      "command:doubleSpacing",
      "command:setSpaceBefore(240)",
      "command:setSpaceAfter(240)",
      "command:increaseIndent",
      "command:decreaseIndent",
      "command:setIndentLeft(720)",
      "command:setIndentRight(720)",
      "command:setIndentFirstLine(360)",
      "command:setIndentFirstLine(hanging)",
      "command:applyStyle(Heading1)",
      "command:applyStyle(Heading2)",
      "command:clearStyle",
      "host:clearFormatting",
    ],
    modes: ["suggesting"],
    kinds: ["reject-mismatch"],
  },
  {
    reason:
      "Suggesting mode does not record the paragraph marks a multi-paragraph paste, Enter over a selection, a cross-paragraph delete or an inserted table of contents create or remove, so rejecting or accepting every change leaves the paragraphs split or unjoined",
    operations: [
      "paste:paragraphs",
      "paste:copied-blocks",
      "key:Enter",
      "key:Delete",
      "command:generateTOC",
    ],
    excludedCases: [{ shape: "tracked-changes", operation: "paste:copied-blocks" }],
    modes: ["suggesting"],
    kinds: ["reject-mismatch", "accept-mismatch"],
  },
  {
    reason:
      "In suggesting mode, typing or inserting over a range deletes the range untracked, and replacing a range that spans paragraphs loses the typed text when the change is accepted",
    operations: REPLACING_OPERATIONS,
    placements: RANGE_PLACEMENTS,
    excludedCases: [{ shape: "tracked-changes", operation: "paste:copied-blocks" }],
    modes: ["suggesting"],
    kinds: ["reject-mismatch", "accept-mismatch"],
    tier: "full",
  },
  {
    reason:
      "Suggesting mode changes table structure (rows, columns, merges, whole tables) directly",
    operations: [
      "command:deleteRow",
      "command:deleteColumn",
      "command:deleteTable",
      "command:addRowAbove",
      "command:addRowBelow",
      "command:addColumnLeft",
      "command:addColumnRight",
      "command:splitCell",
      "command:mergeCells",
    ],
    shapes: ["tables"],
    modes: ["suggesting"],
    kinds: ["reject-mismatch"],
  },

  // -------------------------------------------------------------- other --
  {
    reason:
      "A hyperlink or a character-format change applied across a field drops the field's result text on save, or on rejecting the change",
    operations: ["command:setHyperlink", "command:clearFontSize"],
    shapes: ["fields-links-bookmarks"],
    placements: RANGE_PLACEMENTS,
    kinds: ["readback-blocks", "readback-painted", "reject-mismatch"],
    tier: "full",
  },
];

export const gapCovers = (
  gap: KnownConformanceGap,
  key: ConformanceCaseKey,
  kind: ViolationKind,
  mode: EditorMode,
): boolean =>
  gapApplies(gap, key) &&
  (gap.modes === undefined || gap.modes.includes(mode)) &&
  gap.kinds.includes(kind);

export const gapApplies = (gap: KnownConformanceGap, key: ConformanceCaseKey): boolean =>
  (gap.operations === undefined || gap.operations.includes(key.operation)) &&
  (gap.shapes === undefined || gap.shapes.includes(key.shape)) &&
  (gap.placements === undefined || gap.placements.includes(key.placement)) &&
  !gap.excludedCases?.some(
    ({ shape, operation }) => shape === key.shape && operation === key.operation,
  );
