/**
 * Types for `FormattingBar.vue`, kept in a plain `.ts` module so the package's
 * public `index.ts` can re-export `FormattingBarProps` without resolving a
 * `.vue` SFC. A plain `tsc` (e.g. the Nuxt adapter's typecheck) sees `.vue`
 * imports only through an ambient default-export shim, so named type exports
 * must live in `.ts` files.
 */

import type { ColorValue, Theme } from "@stll/folio-core/types/document";
import type { ListState } from "../utils/listState";
import type { ParagraphAlignment } from "./ui/AlignmentButtons.types";
import type { StyleOption } from "./ui/StylePicker.types";

/**
 * The controls the minimal bar can emit. Mirrors the relevant subset of
 * React's `FormattingAction` (packages/react/src/components/toolbarPrimitives.tsx):
 * mark toggles + list/indent verbs, plus the object-shaped alignment / style /
 * text-color intents.
 */
export type FormattingAction =
  | "bold"
  | "italic"
  | "underline"
  | "bulletList"
  | "numberedList"
  | "indent"
  | "outdent"
  | { type: "alignment"; value: ParagraphAlignment }
  | { type: "applyStyle"; value: string }
  | { type: "textColor"; value: ColorValue | string };

/**
 * Current formatting of the selection driving the bar's active states.
 * Mirrors the fields React's `SelectionFormatting` exposes that the minimal
 * bar reads.
 */
export type SelectionFormatting = {
  /** Whether selected text is bold. */
  bold?: boolean | undefined;
  /** Whether selected text is italic. */
  italic?: boolean | undefined;
  /** Whether selected text is underlined. */
  underline?: boolean | undefined;
  /** Text color (hex, with or without a leading `#`). */
  color?: string | undefined;
  /** Paragraph alignment. */
  alignment?: ParagraphAlignment | undefined;
  /** List state of the current paragraph. */
  listState?: ListState | undefined;
  /** Paragraph style id. */
  styleId?: string | undefined;
  /** Paragraph left indentation in twips. */
  indentLeft?: number | undefined;
};

export type FormattingBarProps = {
  /** Current formatting of the selection (drives active states). */
  currentFormatting?: SelectionFormatting;
  /** Whether the bar is disabled. */
  disabled?: boolean;
  /** Whether undo is available. */
  canUndo?: boolean;
  /** Whether redo is available. */
  canRedo?: boolean;
  /** Additional CSS class name. */
  className?: string;
  /** Whether to show the paragraph-style picker (default: true). */
  showStylePicker?: boolean;
  /** Whether to show the text-color picker (default: true). */
  showTextColorPicker?: boolean;
  /** Whether to show the alignment control (default: true). */
  showAlignmentButtons?: boolean;
  /** Whether to show the list + indent controls (default: true). */
  showListButtons?: boolean;
  /** Document styles for the style picker (Vue `StylePicker` option shape). */
  documentStyles?: StyleOption[];
  /** Document theme — feeds the color picker's theme-color matrix. */
  theme?: Theme | null;
};
