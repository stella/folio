/**
 * Pure list-state helpers used by adapter toolbars to track whether the
 * selection is in a bullet/numbered list and at what indent level.
 *
 * The state itself is core's: one union, read by both adapters, resolved
 * against the numbering definitions rather than guessed from the numbering id.
 * @packageDocumentation
 * @public
 */

import {
  isInListState,
  type ListState,
  listStateLevel,
  type ListType,
  NO_LIST_STATE,
} from "@stll/folio-core/prosemirror";

export type { ListState, ListType };

export function createDefaultListState(): ListState {
  return NO_LIST_STATE;
}

export function createBulletListState(level = 0, numId?: number): ListState {
  return numId === undefined ? { type: "bullet", level } : { type: "bullet", level, numId };
}

export function createNumberedListState(level = 0, numId?: number): ListState {
  return numId === undefined ? { type: "numbered", level } : { type: "numbered", level, numId };
}

export function isBulletListState(state: ListState | undefined): boolean {
  return state?.type === "bullet";
}

export function isNumberedListState(state: ListState | undefined): boolean {
  return state?.type === "numbered";
}

export function isAnyListState(state: ListState | undefined): boolean {
  return isInListState(state);
}

export function getNextIndentLevel(currentLevel: number): number {
  return Math.min(currentLevel + 1, 8);
}

export function getPreviousIndentLevel(currentLevel: number): number {
  return Math.max(currentLevel - 1, 0);
}

export function toggleListType(state: ListState | undefined, targetType: ListType): ListState {
  if (state?.type === targetType) return createDefaultListState();
  const level = listStateLevel(state);
  if (targetType === "bullet") return createBulletListState(level);
  if (targetType === "numbered") return createNumberedListState(level);
  return createDefaultListState();
}
