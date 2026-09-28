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
      "A paragraph a list command numbers states no indentation, while the reopened paragraph reads the level's indentation as its own w:ind",
    operations: [...LIST_TOGGLES, ...LIST_MARKERS, ...LIST_NUMBERING],
    kinds: ["readback-blocks", "readback-painted"],
  },
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
      "Enter at the end of a list item and Backspace at its start take the suggestion-mode paragraph path, not the list path, so suggesting and editing mode produce different lists",
    operations: ["key:Enter", "key:Backspace"],
    shapes: ["single-decimal-list", "single-bullet-list", "outline-level-numbered"],
    placements: ["caret-end", "caret-start"],
    modes: ["suggesting"],
    kinds: ["accept-mismatch"],
  },
  {
    reason:
      "Decreasing a list item's indent to zero, or restyling it, leaves no w:ind, so the numbering level's indent returns on reopen; restyling a paragraph with a direct outline level keeps that level",
    operations: [
      "command:decreaseIndent",
      "command:applyStyle(Heading1)",
      "command:applyStyle(Heading2)",
    ],
    shapes: ["single-decimal-list", "single-bullet-list", "outline-level-numbered"],
    kinds: ["readback-blocks", "readback-painted"],
  },
  {
    reason:
      "Restyling, clearing the style of, or unindenting a range that holds list items or style-numbered headings leaves numbering and indentation in the editor that the saved package does not state",
    operations: [
      "command:decreaseIndent",
      "command:applyStyle(Heading1)",
      "command:applyStyle(Heading2)",
      "command:clearStyle",
      "host:clearFormatting",
    ],
    shapes: ["mixed-lists", "style-numbered-headings", "host-unused-instances"],
    placements: ["cross-paragraph", "document"],
    kinds: ["readback-blocks", "readback-painted", "readback-markdown", "reject-mismatch"],
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
    modes: ["suggesting"],
    kinds: ["reject-mismatch", "accept-mismatch"],
  },
  {
    reason:
      "In suggesting mode, typing or inserting over a range deletes the range untracked, and replacing a range that spans paragraphs loses the typed text when the change is accepted",
    operations: REPLACING_OPERATIONS,
    placements: RANGE_PLACEMENTS,
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
  {
    reason:
      "A suggested deletion inside another author's tracked insertion does not survive a save",
    operations: REPLACING_OPERATIONS,
    shapes: ["tracked-changes", "pending-property-change"],
    modes: ["suggesting"],
    kinds: ["readback-blocks"],
  },

  // -------------------------------------------------------------- other --
  {
    reason:
      "Paragraphs a paste or a structural command creates carry no resolved style attributes (spacing, run defaults) until the document is reopened",
    operations: [
      "paste:paragraphs",
      "paste:copied-blocks",
      "key:Delete",
      "command:deleteColumn",
      "command:addRowAbove",
      "command:addRowBelow",
      "command:addColumnLeft",
      "command:addColumnRight",
      "command:splitCell",
    ],
    kinds: ["readback-blocks", "readback-painted"],
  },
  {
    reason:
      "Adding or deleting a column and adding a row above walk rows without the table map: beside a vertical merge they misplace cells, and they rebuild w:tblGrid from the first row's cell count",
    operations: [
      "command:addColumnLeft",
      "command:addColumnRight",
      "command:deleteColumn",
      "command:addRowAbove",
    ],
    shapes: ["tables"],
    kinds: ["table-grid"],
  },
  {
    reason:
      "Pasting content copied across table rows duplicates paragraph-property source tokens, which the save refuses",
    operations: ["paste:copied-blocks"],
    shapes: ["tables"],
    kinds: ["invalid-model", "table-grid"],
  },
  {
    reason:
      "Replacing a range that spans table rows leaves cells that no longer tile the table grid",
    operations: REPLACING_OPERATIONS,
    shapes: ["tables"],
    placements: ["cross-paragraph"],
    kinds: ["table-grid"],
    tier: "full",
  },
  {
    reason:
      "Replacing a range that spans paragraphs leaves a paragraph without resolved style attributes (spacing, run defaults) until the document is reopened",
    operations: REPLACING_OPERATIONS,
    placements: ["cross-paragraph", "document"],
    kinds: ["readback-blocks", "readback-painted"],
    tier: "full",
  },
  {
    reason:
      "A hyperlink or a character-format change applied across a field drops the field's result text on save, or on rejecting the change",
    operations: ["command:setHyperlink", "command:clearFontSize"],
    shapes: ["fields-links-bookmarks"],
    placements: RANGE_PLACEMENTS,
    kinds: ["readback-blocks", "readback-painted", "reject-mismatch"],
    tier: "full",
  },
  {
    reason:
      "An inserted table of contents reads differently after a save: entry run formatting and page-number runs",
    operations: ["command:generateTOC"],
    shapes: ["outline-level-numbered", "style-numbered-headings"],
    kinds: ["readback-blocks", "readback-painted", "readback-markdown"],
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
  (gap.placements === undefined || gap.placements.includes(key.placement));
