import { GlobalRegistrator } from "@happy-dom/global-registrator";

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { expect, test } from "bun:test";
import { EditorState } from "prosemirror-state";
import { EditorView } from "prosemirror-view";

import { schema } from "@stll/folio-core/prosemirror/schema";

import { useTableResize } from "./useTableResize";

const resizedRowHeight = (height: number, pointerDelta: number): number => {
  const paragraph = schema.nodes.paragraph.create(null, schema.text("before"));
  const cell = schema.nodes.tableCell.createAndFill();
  if (!cell) throw new Error("Expected a table cell");
  const row = schema.nodes.tableRow.create({ height, heightRule: "exact" }, [
    cell,
  ]);
  const table = schema.nodes.table.create({ columnWidths: [3000] }, [row]);
  const doc = schema.nodes.doc.create(null, [paragraph, table]);
  const tablePosition = paragraph.nodeSize;
  const container = document.createElement("div");
  document.body.append(container);
  const view = new EditorView(container, {
    state: EditorState.create({ doc }),
  });
  const layoutTable = document.createElement("div");
  layoutTable.className = "layout-table";
  const layoutRow = document.createElement("div");
  layoutRow.dataset["rowIndex"] = "0";
  layoutTable.append(layoutRow);
  const handle = document.createElement("div");
  handle.className = "layout-table-row-resize-handle";
  handle.dataset["rowIndex"] = "0";
  handle.dataset["tablePmStart"] = String(tablePosition);
  handle.style.top = "50px";
  layoutTable.append(handle);
  container.append(layoutTable);

  const resize = useTableResize();
  let mouseDown: MouseEvent | undefined;
  handle.addEventListener("mousedown", (event) => {
    mouseDown = event;
  });
  handle.dispatchEvent(new MouseEvent("mousedown", { clientY: 50 }));
  if (!mouseDown || !resize.tryStartResize(mouseDown, view))
    throw new Error("Expected the row resize handle to start a drag");

  const uninstall = resize.install();
  document.dispatchEvent(
    new MouseEvent("mousemove", { clientY: 50 + pointerDelta }),
  );
  document.dispatchEvent(
    new MouseEvent("mouseup", { clientY: 50 + pointerDelta }),
  );

  const updatedTable = view.state.doc.nodeAt(tablePosition);
  const updatedRow = updatedTable?.child(0);
  const updatedHeight = updatedRow?.attrs["height"];
  uninstall();
  view.destroy();
  container.remove();
  if (typeof updatedHeight !== "number")
    throw new Error("Expected the resized row to keep an explicit height");
  return updatedHeight;
};

test("dragging a row below the minimum commits the minimum height", () => {
  expect(resizedRowHeight(300, -8)).toBe(200);
});

test("dragging a short spacer row upward preserves its height", () => {
  expect(resizedRowHeight(100, -4)).toBe(100);
});
