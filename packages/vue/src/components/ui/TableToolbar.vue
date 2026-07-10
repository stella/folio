<!--
  Vue table-context toolbar — appears when the cursor is inside a
  table node. Routes UI-level picker values (preset name, hex color,
  "addRowAbove"-style action) into the PM commands registered by
  `TableExtension`. A local `borderSpec` reactive holds the active
  style/width/color so subsequent preset clicks pick them up.

  Composes the five picker children (`TableBorderPicker`,
  `TableBorderColorPicker`, `TableBorderWidthPicker`, `TableCellFillPicker`,
  `TableMoreDropdown` — the last pulling in `TableStyleGallery` /
  `TableGridPicker` / `tableStylePresets`) and routes their emits into the PM
  commands registered by `TableExtension`. The routers stay surfaced via
  `defineExpose` for host-driven use. `TableBorderPreset` / `TableAction` are
  declared locally to mirror the child emit contracts.
-->
<template>
  <template v-if="isInTable">
    <span class="divider" />
    <TableBorderPicker @change="onBorderPreset" />
    <TableBorderColorPicker
      :theme="theme ?? null"
      v-bind="borderColorValueBind"
      @change="onBorderColor"
    />
    <TableBorderWidthPicker @change="onBorderWidth" />
    <TableCellFillPicker :theme="theme ?? null" @change="onCellFill" />
    <TableMoreDropdown
      :can-split="canSplit"
      :can-merge="canMerge"
      :row-count="rowCount"
      :column-count="columnCount"
      :current-justification="currentJustification"
      @action="onMoreAction"
      @cell-margins="onCellMargins"
      @cell-text-direction="onCellTextDirection"
      @row-height="onRowHeight"
    />
  </template>
</template>

<script setup lang="ts">
import { computed, watch } from "vue";
import type { EditorView } from "prosemirror-view";
import type { Command, Transaction } from "prosemirror-state";
import { getTableContext } from "@stll/folio-core/prosemirror/extensions/nodes/TableExtension";
import type { Theme } from "@stll/folio-core/types/document";
import TableBorderPicker from "./TableBorderPicker.vue";
import TableBorderColorPicker from "./TableBorderColorPicker.vue";
import TableBorderWidthPicker from "./TableBorderWidthPicker.vue";
import TableCellFillPicker from "./TableCellFillPicker.vue";
import TableMoreDropdown from "./TableMoreDropdown.vue";

/**
 * A toolbar command factory: called with whatever arguments the specific
 * command needs, it returns a ProseMirror `Command`.
 */
type CommandFactory = (...args: readonly unknown[]) => Command;

/** Border preset `TableBorderPicker` emits. */
type TableBorderPreset =
  | "all"
  | "none"
  | "box"
  | "inside"
  | "insideH"
  | "insideV"
  | "top"
  | "bottom"
  | "left"
  | "right";

/**
 * Action string `TableMoreDropdown` emits. Kept as a widened
 * string so unmapped actions fall through to `exec(action)` directly, matching
 * upstream behaviour.
 */
type TableAction = string;

const props = defineProps<{
  view: EditorView | null;
  getCommands: () => Record<string, CommandFactory>;
  stateTick: number;
  /** Document theme — fed through to the border/fill color pickers. */
  theme?: Theme | null;
}>();

// Single source of truth for the table-context state the toolbar reads —
// recomputed on every editor transaction. `isInTable` and the "more"
// dropdown's gating/active state all derive from this one walk.
const tableCtx = computed(() => {
  void props.stateTick;
  const v = props.view;
  return v ? getTableContext(v.state) : null;
});

const isInTable = computed(() => !!tableCtx.value?.isInTable);

// "Split cell" needs the current cell to span more than one row/column —
// stricter than `getTableContext`'s `canSplitCell` ("just in a cell"),
// so it reads the cell node's colspan/rowspan directly.
const canSplit = computed(() => {
  void props.stateTick;
  const v = props.view;
  if (!v) return false;
  const { $from } = v.state.selection;
  for (let depth = $from.depth; depth > 0; depth--) {
    if ($from.node(depth).type.name === "tableCell") {
      const cell = $from.node(depth);
      const colspanRaw = cell.attrs["colspan"];
      const rowspanRaw = cell.attrs["rowspan"];
      const colspan = typeof colspanRaw === "number" ? colspanRaw : 1;
      const rowspan = typeof rowspanRaw === "number" ? rowspanRaw : 1;
      return colspan > 1 || rowspan > 1;
    }
  }
  return false;
});

// Current cell border color (RGB hex, no #) for the color picker's active
// swatch. Only the literal-rgb ColorValue shape resolves without theme lookup.
const currentBorderColorHex = computed<string | undefined>(() => {
  const c = tableCtx.value?.cellBorderColor;
  if (c && typeof c === "object" && "rgb" in c && typeof c.rgb === "string") {
    return c.rgb;
  }
  return undefined;
});

// Omit `value` entirely (not `undefined`) under exactOptionalPropertyTypes so
// an absent border color drops the key rather than passing an explicit undefined.
const borderColorValueBind = computed<{ value?: string }>(() =>
  currentBorderColorHex.value !== undefined ? { value: currentBorderColorHex.value } : {},
);

const canMerge = computed(() => !!tableCtx.value?.hasMultiCellSelection);
const rowCount = computed(() => tableCtx.value?.rowCount ?? 0);
const columnCount = computed(() => tableCtx.value?.columnCount ?? 0);
const currentJustification = computed<"left" | "center" | "right">(() => {
  const j = tableCtx.value?.table?.attrs["justification"];
  return j === "center" || j === "right" ? j : "left";
});

// Default style/size/color match TableExtension's solid-border defaults.
// Read in callbacks only — no template/computed consumer, so a plain
// object is sufficient.
const borderSpec: { style: string; size: number; color: { rgb: string } } = {
  style: "single",
  size: 4,
  color: { rgb: "000000" },
};

// Sync the spec's color with the current cell's existing border color
// so the first "All borders" / "Outside" click on a pre-bordered table
// preserves the visible color instead of stamping black over it.
watch(
  [() => props.view, () => props.stateTick],
  () => {
    const v = props.view;
    if (!v) return;
    const ctx = getTableContext(v.state);
    const c = ctx.cellBorderColor;
    if (!c || !ctx.isInTable) return;
    // `cellBorderColor` is a ColorValue — only the literal rgb shape
    // can flow into the OOXML border spec without theme resolution.
    if (typeof c === "object" && "rgb" in c && typeof c.rgb === "string") {
      borderSpec.color = { rgb: c.rgb };
    }
  },
  { immediate: true },
);

function exec(name: string, ...args: unknown[]): boolean {
  const v = props.view;
  if (!v) return false;
  const factory = props.getCommands()[name];
  if (!factory) return false;
  const command = factory(...args);
  command(v.state, (tr: Transaction) => v.dispatch(tr), v);
  v.focus();
  return true;
}

function onBorderPreset(preset: TableBorderPreset) {
  // Snapshot so the spec stored in PM node attrs doesn't alias our
  // local mutable object.
  const spec = { ...borderSpec, color: { ...borderSpec.color } };
  switch (preset) {
    case "all":
      exec("setAllTableBorders", spec);
      return;
    case "none":
      exec("removeTableBorders");
      return;
    case "box":
      exec("setOutsideTableBorders", spec);
      return;
    case "inside":
    case "insideH":
    case "insideV":
      // BorderPreset has no insideH/insideV split; both fall back to
      // the umbrella "inside" preset.
      exec("setInsideTableBorders", spec);
      return;
    case "top":
    case "bottom":
    case "left":
    case "right":
      exec("setCellBorder", preset, spec, true);
      return;
  }
}

function onBorderColor(hex: string) {
  borderSpec.color = { rgb: hex.replace(/^#/, "") };
  exec("setTableBorderColor", hex);
}

function onBorderWidth(eighths: number) {
  borderSpec.size = eighths;
  exec("setTableBorderWidth", eighths);
}

function onCellFill(hex: string) {
  exec("setCellFillColor", hex);
}

// Actions whose menu name diverges from the core command name, plus
// their extra args. Everything not in this map falls through to
// `exec(action)` directly.
const moreActionMap: Partial<Record<TableAction, [string, ...unknown[]]>> = {
  autoFit: ["autoFitContents"],
  alignTableLeft: ["setTableProperties", { justification: "left" }],
  alignTableCenter: ["setTableProperties", { justification: "center" }],
  alignTableRight: ["setTableProperties", { justification: "right" }],
  verticalAlignTop: ["setCellVerticalAlign", "top"],
  verticalAlignMiddle: ["setCellVerticalAlign", "center"],
  verticalAlignBottom: ["setCellVerticalAlign", "bottom"],
};

function onMoreAction(action: TableAction) {
  // Dialog action is a v1.x followup — explicit no-op so the menu
  // closes cleanly without dispatching a phantom command.
  if (action === "tableProperties") return;
  const mapped = moreActionMap[action];
  if (mapped) exec(...mapped);
  else exec(action);
}

function onCellMargins(margins: { top?: number; bottom?: number; left?: number; right?: number }) {
  exec("setCellMargins", margins);
}

function onCellTextDirection(direction: string | null) {
  exec("setCellTextDirection", direction);
}

function onRowHeight(value: { height: number | null; rule?: "auto" | "atLeast" | "exact" }) {
  exec("setRowHeight", value.height, value.rule);
}

// Surface the command routers + gating state until the picker children are
// ported and can bind these directly in the template.
defineExpose({
  isInTable,
  canSplit,
  canMerge,
  rowCount,
  columnCount,
  currentJustification,
  onBorderPreset,
  onBorderColor,
  onBorderWidth,
  onCellFill,
  onMoreAction,
  onCellMargins,
  onCellTextDirection,
  onRowHeight,
});
</script>

<style scoped>
.divider {
  width: 1px;
  height: 20px;
  margin: 0 6px;
  background: var(--doc-border);
  flex-shrink: 0;
}
</style>
