import { describe, expect, test } from "bun:test";

import { resetImeCaretAnchor, syncImeCaretAnchor } from "./imeCaretAnchor";

type MockEditorViewOptions = {
  focused?: boolean;
  composing?: boolean;
  empty?: boolean;
  head?: number;
  coords?: { left: number; top: number };
  throws?: boolean;
};

const mockEditorView = ({
  focused = true,
  composing = false,
  empty = true,
  head = 5,
  coords = { left: -9999, top: 0 },
  throws = false,
}: MockEditorViewOptions = {}) => ({
  composing,
  hasFocus: () => focused,
  state: { selection: { empty, head } },
  coordsAtPos: () => {
    if (throws) {
      throw new RangeError("Stale editor selection");
    }
    return coords;
  },
});

const createHost = (transform = "") => ({ style: { transform } });

describe("syncImeCaretAnchor", () => {
  test("translates the hidden caret onto the painted caret", () => {
    const host = createHost();

    const anchored = syncImeCaretAnchor({
      hiddenHost: host,
      editorView: mockEditorView(),
      visibleCaret: { left: 320, top: 240 },
    });

    expect(anchored).toBe(true);
    expect(host.style.transform).toBe("translate3d(10319px, 240px, 0)");
  });

  test("measures from the untransformed baseline on repeated updates", () => {
    const host = createHost("translate3d(10319px, 240px, 0)");
    const editorView = mockEditorView();
    editorView.coordsAtPos = () => {
      expect(host.style.transform).toBe("");
      return { left: -9999, top: 0 };
    };

    syncImeCaretAnchor({
      hiddenHost: host,
      editorView,
      visibleCaret: { left: 321, top: 241 },
    });

    expect(host.style.transform).toBe("translate3d(10320px, 241px, 0)");
  });

  test("resets the anchor for range selections and unfocused editors", () => {
    const host = createHost("translate3d(10px, 20px, 0)");

    expect(
      syncImeCaretAnchor({
        hiddenHost: host,
        editorView: mockEditorView({ empty: false }),
        visibleCaret: { left: 10, top: 20 },
      }),
    ).toBe(false);
    expect(host.style.transform).toBe("");

    host.style.transform = "translate3d(10px, 20px, 0)";
    expect(
      syncImeCaretAnchor({
        hiddenHost: host,
        editorView: mockEditorView({ focused: false }),
        visibleCaret: { left: 10, top: 20 },
      }),
    ).toBe(false);
    expect(host.style.transform).toBe("");
  });

  test("preserves the anchor while composition is active", () => {
    const host = createHost("translate3d(10px, 20px, 0)");

    expect(
      syncImeCaretAnchor({
        hiddenHost: host,
        editorView: mockEditorView({ composing: true }),
        visibleCaret: null,
      }),
    ).toBe(false);
    expect(host.style.transform).toBe("translate3d(10px, 20px, 0)");

    resetImeCaretAnchor(host);
    expect(host.style.transform).toBe("");
  });

  test("restores the previous anchor when native caret measurement fails", () => {
    const host = createHost("translate3d(10px, 20px, 0)");

    expect(
      syncImeCaretAnchor({
        hiddenHost: host,
        editorView: mockEditorView({ throws: true }),
        visibleCaret: { left: 320, top: 240 },
      }),
    ).toBe(false);
    expect(host.style.transform).toBe("translate3d(10px, 20px, 0)");
  });
});
