/**
 * Known gaps in the editor-command conformance matrix.
 *
 * Each entry names the cases it covers and the violation kinds it expects.
 * A covered case that shows an expected violation passes; a violation no
 * entry covers fails. An entry that covered cases in a run but saw none of
 * its violations is stale and fails too, so a fix removes its entry in the
 * same change — entries tied to an issue flip that way when the fix lands.
 */

import { panic } from "better-result";

import type { ViolationKind } from "./editorCommandConformance";
import { LIST_PASTE_RESOLUTION_KEYS } from "./editorCommandConformance.listPaste";
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
  /** Strict coverage can apply to a complete pair or one exact selection case. */
  excludedCases?: readonly (
    | ({ scope: "shapeOperation" } & Pick<ConformanceCaseKey, "shape" | "operation">)
    | ({ scope: "case" } & ConformanceCaseKey)
  )[];
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
] as const;

const RANGE_PLACEMENTS: readonly SelectionPlacement[] = [
  "word",
  "paragraph",
  "cross-paragraph",
  "document",
];

// The reported #1404 image/node cases, independent of full-tier placement expansion.
export const LEGACY_NODE_REPLACEMENTS = [
  { operation: "command:insertHyperlink", status: "strict" },
  { operation: "command:insertFootnote", status: "strict" },
  { operation: "command:insertEndnote", status: "strict" },
  { operation: "key:Shift-Enter", status: "strict" },
  { operation: "type:text", status: "strict" },
  { operation: "type:text(mid)", status: "strict" },
  { operation: "type:number-marker", status: "strict" },
  { operation: "paste:plain", status: "strict" },
  { operation: "paste:table", status: "strict" },
  { operation: "paste:list", status: "strict" },
] as const satisfies readonly (
  | { operation: string; status: "strict" }
  | { operation: string; status: "expectedFailure"; kind: ViolationKind }
)[];

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
    excludedCases: LEGACY_NODE_REPLACEMENTS.map(({ operation }) => ({
      scope: "case" as const,
      shape: "image",
      operation,
      placement: "node" as const,
    })),
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
    excludedCases: [
      ...LIST_PASTE_RESOLUTION_KEYS.map(({ shape, operation }) => ({
        scope: "shapeOperation" as const,
        shape,
        operation,
      })),
      { scope: "shapeOperation", shape: "tracked-changes", operation: "paste:copied-blocks" },
      { scope: "shapeOperation", shape: "tables", operation: "paste:copied-blocks" },
    ],
    modes: ["suggesting"],
    kinds: ["reject-mismatch", "accept-mismatch"],
  },
  {
    reason:
      "In suggesting mode, typing or inserting over a range deletes the range untracked, and replacing a range that spans paragraphs loses the typed text when the change is accepted",
    operations: REPLACING_OPERATIONS,
    placements: RANGE_PLACEMENTS,
    excludedCases: [
      { scope: "shapeOperation", shape: "tracked-changes", operation: "paste:copied-blocks" },
    ],
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
  !gap.excludedCases?.some((excluded) => {
    if (excluded.shape !== key.shape || excluded.operation !== key.operation) return false;
    switch (excluded.scope) {
      case "shapeOperation":
        return true;
      case "case":
        return excluded.placement === key.placement;
      default: {
        const impossible: never = excluded;
        return panic(`Unknown conformance gap exclusion: ${String(impossible)}`);
      }
    }
  });
