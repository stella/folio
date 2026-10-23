import { commitColumnResize, commitRightEdgeResize, commitRowResize } from "@stll/folio-core/prosemirror/tableResize";

export const handleMouseUp = () => {
  commitColumnResize(view, pair);
  commitRightEdgeResize(view, edge);
  commitRowResize(view, row);
};

export const otherTableCommand = () => transaction.setNodeMarkup(position, undefined, { columnWidths });
