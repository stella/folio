import type { EditorState } from "prosemirror-state";
import {
  getCaretPositionFromDom,
  clickToPositionDom,
} from "../../packages/core/src/layout-bridge/dom/clickToPositionDom";
import { resolvePainterTarget } from "./painterTargetCommit";
import type { BrowserDragTarget } from "./browserDragTarget";

type PaintedTargetEditor = {
  getEditorRef: () => { getView: () => { state: EditorState } | null } | null;
  onLayoutChange: (listener: () => void) => () => void;
};

export const resolvePaintedTableTarget = (ref: PaintedTargetEditor) =>
  resolvePainterTarget({
    subscribe: ref.onLayoutChange,
    read: () => {
      const view = ref.getEditorRef()?.getView();
      if (!view) throw new Error("browser editor unavailable");
      const positions: { pos: number; paragraphPos: number }[] = [];
      view.state.doc.descendants((node, pos) => {
        if (positions.length >= 2) return false;
        if (node.type.name !== "tableCell") return true;
        let paragraphPos: number | null = null;
        node.descendants((child, offset) => {
          if (paragraphPos !== null) return false;
          if (child.isTextblock) paragraphPos = pos + 1 + offset;
          return !child.isTextblock;
        });
        if (paragraphPos !== null) positions.push({ pos, paragraphPos });
        return false;
      });

      const [anchor, head] = positions;
      if (!anchor || !head) return { type: "absent" } as const;
      const cell = (position: number) =>
        document.querySelector<HTMLElement>(
          `.layout-page-content .layout-table-cell[data-pm-start="${position}"]`,
        );
      const anchorCell = cell(anchor.paragraphPos);
      const headCell = cell(head.paragraphPos);
      if (!anchorCell || !headCell) return null;
      anchorCell.scrollIntoView({ block: "nearest" });
      headCell.scrollIntoView({ block: "nearest" });
      // Scrolling can replace virtualized paint; resolve the nodes again in this read.
      const from = cell(anchor.paragraphPos)?.getBoundingClientRect();
      const to = cell(head.paragraphPos)?.getBoundingClientRect();
      if (!from || !to || from.width === 0 || to.width === 0) return null;
      return {
        type: "ready",
        anchor: anchor.pos,
        head: head.pos,
        from: { x: from.x + from.width / 2, y: from.y + from.height / 2 },
        to: { x: to.x + to.width / 2, y: to.y + to.height / 2 },
      } as const;
    },
  });

export const resolvePaintedDragTarget = (
  ref: PaintedTargetEditor,
  wanted: Exclude<BrowserDragTarget, "table">,
) =>
  resolvePainterTarget({
    subscribe: ref.onLayoutChange,
    read: () => {
      const view = ref.getEditorRef()?.getView();
      if (!view) throw new Error("browser editor unavailable");
      const targets: { pos: number; size: number; type: "paragraph" | "inline" }[] = [];
      const targetName = {
        list: "paragraph",
        note: "footnoteRef",
        field: "field",
        inlineObject: "image",
      }[wanted];
      view.state.doc.descendants((node, pos) => {
        const matches =
          node.type.name === targetName &&
          (wanted !== "list" ||
            (node.attrs["numPr"] !== null && node.attrs["numPr"] !== undefined));
        if (targets.length >= 2) return false;
        if (
          matches ||
          (wanted === "note" && node.marks.some((mark) => mark.type.name === targetName))
        ) {
          targets.push({
            pos,
            size: node.nodeSize,
            type: node.type.name === "paragraph" ? "paragraph" : "inline",
          });
        }
        return true;
      });
      const first = targets.at(0);
      if (!first) return { type: "absent" } as const;
      const last = targets.at(1) ?? first;
      // Paragraph node boundaries are not caret positions. Start inside the
      // paragraph, or in text before an atom so image mousedown can drag text.
      const rawFrom = first.type === "paragraph" ? first.pos + 1 : first.pos - 1;
      const rawTo =
        last.type === "paragraph"
          ? last.pos + Math.min(2, last.size - 1)
          : last.pos + last.size + 2;
      const from = Math.max(view.state.doc.resolve(first.pos).start(), rawFrom);
      const to = Math.min(
        view.state.doc.content.size - 1,
        last.type === "paragraph"
          ? last.pos + last.size - 1
          : view.state.doc.resolve(last.pos).end(),
        rawTo,
      );
      const positions = {
        from,
        to,
        targetFrom: first.pos,
        targetTo: first.pos + first.size,
        targetType: first.type,
      };
      const coordinate = (position: number) => {
        const caret = getCaretPositionFromDom(document.body, position, new DOMRect());
        if (!caret) return null;
        const y = caret.y + caret.height / 2;
        for (const offset of [0, 0.25, -0.25, 0.5, -0.5]) {
          const x = caret.x + offset;
          if (clickToPositionDom(document.body, x, y) === position) return { x, y };
        }
        return null;
      };
      const fromCoordinate = coordinate(positions.from);
      const toCoordinate = coordinate(positions.to);
      return fromCoordinate && toCoordinate
        ? ({ type: "ready", positions, from: fromCoordinate, to: toCoordinate } as const)
        : null;
    },
  });
