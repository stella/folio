/**
 * Text Formatting Commands — thin re-exports from extension system
 *
 * Toggle marks, set marks, clear formatting, hyperlinks.
 * All implementations live in extensions/marks/; this file re-exports
 * for backward compatibility.
 */

import type { Mark } from "prosemirror-model";
import type { Command } from "prosemirror-state";

import type { TextFormatting } from "../../types/document";
import { textFormattingToMarks as _textFormattingToMarks } from "../extensions/marks/markUtils";
import { singletonManager, schema } from "../schema";
import type { TextColorAttrs } from "../schema";

// Utility re-exports from markUtils (used by toolbar, conversion, etc.)
export {
  isMarkActive,
  getMarkAttr,
  clearFormatting,
  createSetMarkCommand,
  createRemoveMarkCommand,
} from "../extensions/marks/markUtils";

// Hyperlink query helpers (used by toolbar)
export {
  isHyperlinkActive,
  getHyperlinkAttrs,
  getSelectedText,
} from "../extensions/marks/HyperlinkExtension";

// ============================================================================
// PARAGRAPH DEFAULT FORMATTING HELPERS
// ============================================================================

/**
 * textFormattingToMarks — wraps markUtils version to use singleton schema
 */

export function textFormattingToMarks(formatting: TextFormatting): Mark[] {
  return _textFormattingToMarks(formatting, schema);
}

// ============================================================================
// COMMANDS — delegated to singleton extension manager
// ============================================================================

// The singleton validates the complete built-in registry during startup.
const cmds = singletonManager;

// Toggle marks (simple on/off)
export const toggleBold: Command = cmds.requireCommand("toggleBold")();
export const toggleItalic: Command = cmds.requireCommand("toggleItalic")();
export const toggleUnderline: Command = cmds.requireCommand("toggleUnderline")();
export const toggleStrike: Command = cmds.requireCommand("toggleStrike")();
export const toggleSuperscript: Command = cmds.requireCommand("toggleSuperscript")();
export const toggleSubscript: Command = cmds.requireCommand("toggleSubscript")();

// Set marks (with attributes)
export function setTextColor(attrs: TextColorAttrs): Command {
  return cmds.requireCommand("setTextColor")(attrs);
}
export const clearTextColor: Command = cmds.requireCommand("clearTextColor")();

export function setHighlight(color: string): Command {
  return cmds.requireCommand("setHighlight")(color);
}
export const clearHighlight: Command = cmds.requireCommand("clearHighlight")();

export function setFontSize(size: number): Command {
  return cmds.requireCommand("setFontSize")(size);
}
export const clearFontSize: Command = cmds.requireCommand("clearFontSize")();

export function setFontFamily(fontName: string): Command {
  return cmds.requireCommand("setFontFamily")(fontName);
}
export const clearFontFamily: Command = cmds.requireCommand("clearFontFamily")();

export function setUnderlineStyle(style: string, color?: TextColorAttrs): Command {
  return cmds.requireCommand("setUnderlineStyle")(style, color);
}

// Hyperlink commands
export function setHyperlink(href: string, tooltip?: string): Command {
  return cmds.requireCommand("setHyperlink")(href, tooltip);
}
export const removeHyperlink: Command = cmds.requireCommand("removeHyperlink")();

export function insertHyperlink(text: string, href: string, tooltip?: string): Command {
  return cmds.requireCommand("insertHyperlink")(text, href, tooltip);
}
