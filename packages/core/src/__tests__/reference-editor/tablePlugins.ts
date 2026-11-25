import { Plugin, PluginKey } from "prosemirror-state";
import {
  CellSelection,
  columnResizing,
  tableEditing,
} from "prosemirror-tables";
import { Decoration, DecorationSet } from "prosemirror-view";

import { pasteTableCells } from "../../prosemirror/tableCellPaste";
import { tablePositionedBookmarksPlugin } from "./tablePositionedBookmarks";

export const createReferenceTablePlugins = (): Plugin[] => {
  // Active cell highlight plugin — adds a CSS class to the cell containing the cursor
  const activeCellKey = new PluginKey("activeCell");
  const activeCellPlugin = new Plugin({
    key: activeCellKey,
    props: {
      decorations(state) {
        const { selection } = state;
        // Skip if already a CellSelection (prosemirror-tables handles that)
        if (selection instanceof CellSelection) {
          return DecorationSet.empty;
        }

        const { $from } = selection;
        for (let d = $from.depth; d > 0; d--) {
          const node = $from.node(d);
          if (
            node.type.name === "tableCell" ||
            node.type.name === "tableHeader"
          ) {
            const pos = $from.before(d);
            return DecorationSet.create(state.doc, [
              Decoration.node(pos, pos + node.nodeSize, {
                class: "activeCell",
              }),
            ]);
          }
        }
        return DecorationSet.empty;
      },
    },
  });

  // Cell paste runs ahead of `tableEditing`, whose own cell paste splits
  // merges at positions that ignore where the table starts.
  const tableCellPastePlugin = new Plugin({
    key: new PluginKey("tableCellPaste"),
    props: {
      handlePaste: (view, _event, slice) =>
        pasteTableCells(view.state, slice, view.dispatch),
    },
  });

  return [
    columnResizing({
      handleWidth: 5,
      cellMinWidth: 25,
      lastColumnResizable: true,
    }),
    tableCellPastePlugin,
    tableEditing(),
    tablePositionedBookmarksPlugin(),
    activeCellPlugin,
  ];
};
