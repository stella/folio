import type {
  CaretPosition,
  SelectionRect,
} from "@stll/folio-core/layout-bridge/engine/selectionRects";

export type RemoteSelectionGeometry = {
  caretPosition: CaretPosition | null;
  selectionRects: SelectionRect[];
};

const hasSameCaretPosition = (previous: CaretPosition | null, next: CaretPosition | null) => {
  if (previous === null || next === null) {
    return previous === next;
  }

  return (
    previous.height === next.height &&
    previous.pageIndex === next.pageIndex &&
    previous.x === next.x &&
    previous.y === next.y
  );
};

const hasSameSelectionRect = (previous: SelectionRect, next: SelectionRect) =>
  previous.height === next.height &&
  previous.pageIndex === next.pageIndex &&
  previous.width === next.width &&
  previous.x === next.x &&
  previous.y === next.y;

export const retainRemoteSelectionGeometry = (
  previous: RemoteSelectionGeometry | null,
  next: RemoteSelectionGeometry,
) => {
  if (
    previous === null ||
    !hasSameCaretPosition(previous.caretPosition, next.caretPosition) ||
    previous.selectionRects.length !== next.selectionRects.length
  ) {
    return next;
  }

  return previous.selectionRects.every((rect, index) => {
    const nextRect = next.selectionRects.at(index);
    return nextRect !== undefined && hasSameSelectionRect(rect, nextRect);
  })
    ? previous
    : next;
};
