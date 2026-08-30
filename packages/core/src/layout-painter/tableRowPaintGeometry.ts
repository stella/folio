const CSS_PIXEL_ROUNDING_EPSILON = 1e-6;

type OwnedRowBottomBorderOffsetsOptions = {
  origin: number;
  rowHeights: readonly number[];
  /** Whether the row ending at this index owns a visible shared bottom edge. */
  snapAfterRow: readonly boolean[];
};

/**
 * Offset bottom-owned shared edges onto CSS pixel boundaries. Source row
 * geometry remains unchanged, so content floors and the final band edge stay
 * exact; only the independently painted border moves.
 */
export const ownedRowBottomBorderOffsets = ({
  origin,
  rowHeights,
  snapAfterRow,
}: OwnedRowBottomBorderOffsetsOptions): number[] => {
  const offsets: number[] = [];
  let boundary = origin;
  for (let rowIndex = 0; rowIndex < rowHeights.length; rowIndex++) {
    boundary += rowHeights[rowIndex] ?? 0;
    if (rowIndex === rowHeights.length - 1 || snapAfterRow[rowIndex] !== true) {
      offsets.push(0);
      continue;
    }
    offsets.push(Math.ceil(boundary - CSS_PIXEL_ROUNDING_EPSILON) - boundary);
  }
  return offsets;
};
