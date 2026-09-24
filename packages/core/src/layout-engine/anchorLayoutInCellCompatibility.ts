/**
 * First `compatibilityMode` that ignores `wp:anchor/@layoutInCell="0"`.
 *
 * From `compatibilityMode` 15 onward, an anchor inside a table cell is laid
 * out in the cell regardless of what `wp:anchor/@layoutInCell` states: the
 * attribute stops opting the anchor out of the cell's coordinate space.
 * Earlier modes still honour an explicit `layoutInCell="0"` and let the
 * anchor position itself relative to the page. A document that declares no
 * `compatSetting` at all is read with the oldest semantics, so it honours
 * the attribute too.
 */
const FIRST_ALWAYS_IN_CELL_COMPATIBILITY_MODE = 15;

/**
 * `true` when the document's compatibility mode forces every table-cell
 * anchor to lay out inside its cell, ignoring an authored
 * `layoutInCell="0"`. `false` means the authored attribute (or its
 * OOXML default of "confined to the cell") should be honoured as stated.
 */
export const resolveAnchorLayoutInCellCompatibility = (
  compatibilityMode: number | undefined,
): boolean =>
  compatibilityMode !== undefined && compatibilityMode >= FIRST_ALWAYS_IN_CELL_COMPATIBILITY_MODE;
