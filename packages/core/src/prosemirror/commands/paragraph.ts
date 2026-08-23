/**
 * Paragraph Formatting Commands — thin re-exports from extension system
 *
 * Alignment, line spacing, indentation, lists, paragraph styles.
 * All implementations live in extensions/; this file re-exports
 * for backward compatibility.
 */

import type { Command } from "prosemirror-state";

import type {
  ParagraphAlignment,
  LineSpacingRule,
  TabStopAlignment,
  TabLeader,
} from "../../types/document";
// Re-export types and query helpers from extensions
import type { ResolvedStyleAttrs } from "../extensions/core/ParagraphExtension";
import { singletonManager } from "../schema";

export type { ResolvedStyleAttrs } from "../extensions/core/ParagraphExtension";
export {
  getParagraphAlignment,
  getStyleId,
  getParagraphTabs,
  getParagraphBidi,
} from "../extensions/core/ParagraphExtension";
export { isInList, getListInfo } from "../extensions/features/ListExtension";

// ============================================================================
// COMMANDS — delegated to singleton extension manager
// ============================================================================

// The singleton validates the complete built-in registry during startup.
const cmds = singletonManager;

// Alignment
export function setAlignment(alignment: ParagraphAlignment): Command {
  return cmds.requireCommand("setAlignment")(alignment);
}
export const alignLeft: Command = cmds.requireCommand("alignLeft")();
export const alignCenter: Command = cmds.requireCommand("alignCenter")();
export const alignRight: Command = cmds.requireCommand("alignRight")();
export const alignJustify: Command = cmds.requireCommand("alignJustify")();

// Line spacing
export function setLineSpacing(value: number, rule: LineSpacingRule = "auto"): Command {
  return cmds.requireCommand("setLineSpacing")(value, rule);
}
export const singleSpacing: Command = cmds.requireCommand("singleSpacing")();
export const oneAndHalfSpacing: Command = cmds.requireCommand("oneAndHalfSpacing")();
export const doubleSpacing: Command = cmds.requireCommand("doubleSpacing")();

// Indentation
export function increaseIndent(amount: number = 720): Command {
  return cmds.requireCommand("increaseIndent")(amount);
}
export function decreaseIndent(amount: number = 720): Command {
  return cmds.requireCommand("decreaseIndent")(amount);
}
export function setIndentLeft(twips: number): Command {
  return cmds.requireCommand("setIndentLeft")(twips);
}
export function setIndentRight(twips: number): Command {
  return cmds.requireCommand("setIndentRight")(twips);
}
export function setIndentFirstLine(twips: number, hanging?: boolean): Command {
  return cmds.requireCommand("setIndentFirstLine")(twips, hanging);
}

// Lists
export const toggleBulletList: Command = cmds.requireCommand("toggleBulletList")();
export const toggleNumberedList: Command = cmds.requireCommand("toggleNumberedList")();
export const increaseListLevel: Command = cmds.requireCommand("increaseListLevel")();
export const decreaseListLevel: Command = cmds.requireCommand("decreaseListLevel")();
export const removeList: Command = cmds.requireCommand("removeList")();

// Spacing
export function setSpaceBefore(twips: number): Command {
  return cmds.requireCommand("setSpaceBefore")(twips);
}
export function setSpaceAfter(twips: number): Command {
  return cmds.requireCommand("setSpaceAfter")(twips);
}

// Paragraph styles
export function applyStyle(styleId: string, resolvedAttrs?: ResolvedStyleAttrs): Command {
  return cmds.requireCommand("applyStyle")(styleId, resolvedAttrs);
}
export const clearStyle: Command = cmds.requireCommand("clearStyle")();

// Section breaks
export function insertSectionBreak(
  breakType: "nextPage" | "continuous" | "oddPage" | "evenPage",
): Command {
  return cmds.requireCommand("insertSectionBreak")(breakType);
}
export const removeSectionBreak: Command = cmds.requireCommand("removeSectionBreak")();

// Tab stops
export function addTabStop(
  position: number,
  alignment: TabStopAlignment = "left",
  leader: TabLeader = "none",
): Command {
  return cmds.requireCommand("addTabStop")(position, alignment, leader);
}
export function removeTabStop(position: number): Command {
  return cmds.requireCommand("removeTabStop")(position);
}

// Text direction
export const toggleBidi: Command = cmds.requireCommand("toggleBidi")();
export const setRtl: Command = cmds.requireCommand("setRtl")();
export const setLtr: Command = cmds.requireCommand("setLtr")();

// Table of Contents
export const generateTOC: Command = cmds.requireCommand("generateTOC")();
