import { describe, expect, test } from "bun:test";

import { retainRemoteSelectionGeometry } from "./remoteSelectionGeometry";

const geometry = () => ({
  caretPosition: { height: 18, pageIndex: 0, x: 24, y: 48 },
  selectionRects: [
    { height: 18, pageIndex: 0, width: 72, x: 24, y: 48 },
    { height: 18, pageIndex: 0, width: 32, x: 24, y: 66 },
  ],
});

describe("retainRemoteSelectionGeometry", () => {
  test("reaches a fixed point for equivalent recomputations", () => {
    const previous = geometry();
    const retained = retainRemoteSelectionGeometry(previous, geometry());

    expect(retained).toBe(previous);
    expect(retainRemoteSelectionGeometry(retained, geometry())).toBe(previous);
  });

  test("adopts changed caret and selection geometry", () => {
    const previous = geometry();
    const movedCaret = geometry();
    movedCaret.caretPosition.x += 1;
    const resizedSelection = geometry();
    const firstRect = resizedSelection.selectionRects.at(0);
    if (firstRect === undefined) {
      throw new Error("Expected selection rectangle fixture.");
    }
    firstRect.width += 1;

    expect(retainRemoteSelectionGeometry(previous, movedCaret)).toBe(movedCaret);
    expect(retainRemoteSelectionGeometry(previous, resizedSelection)).toBe(resizedSelection);
  });
});
