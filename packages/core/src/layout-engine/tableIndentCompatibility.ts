import type { TableBlock } from "./types";

/**
 * First `compatibilityMode` that measures `w:tblInd` to the table border.
 *
 * Word 2013 moved the reference edge: from that mode on, `w:tblInd` offsets
 * the border and the leading cell margin pushes the text further in. Earlier
 * modes measure the indent to the leading cell's text edge and pull the border
 * back by that cell's leading margin, so the first character lands exactly on
 * the indent. A document that declares no `compatSetting` at all is read with
 * the oldest semantics, so it takes the legacy rule too.
 */
const FIRST_BORDER_EDGE_INDENT_COMPATIBILITY_MODE = 15;

export const resolveTableIndentCompatibility = (
  compatibilityMode: number | undefined,
): NonNullable<TableBlock["indentCompatibility"]> | undefined =>
  compatibilityMode !== undefined &&
  compatibilityMode >= FIRST_BORDER_EDGE_INDENT_COMPATIBILITY_MODE
    ? undefined
    : { type: "legacy" };
