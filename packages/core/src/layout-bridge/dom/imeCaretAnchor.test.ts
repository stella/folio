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

  test("measures with the current translation in place and takes it off", () => {
    const host = createHost("translate3d(10319px, 240px, 0)");
    const editorView = mockEditorView();
    let measuredTransform: string | undefined;
    editorView.coordsAtPos = () => {
      measuredTransform = host.style.transform;
      // The hidden caret stands where the current translation put it.
      return { left: -9999 + 10319, top: 240 };
    };

    syncImeCaretAnchor({
      hiddenHost: host,
      editorView,
      visibleCaret: { left: 321, top: 241 },
    });

    expect(measuredTransform).toBe("translate3d(10319px, 240px, 0)");
    expect(host.style.transform).toBe("translate3d(10320px, 241px, 0)");
  });

  test("does not write the transform when the offset is unchanged", () => {
    let writes = 0;
    let transform = "translate3d(10319px, 240px, 0px)";
    const host = {
      style: {
        get transform() {
          return transform;
        },
        set transform(value: string) {
          writes += 1;
          transform = value;
        },
      },
    };
    const editorView = mockEditorView();
    editorView.coordsAtPos = () => ({ left: -9999 + 10319, top: 240 });

    syncImeCaretAnchor({
      hiddenHost: host,
      editorView,
      visibleCaret: { left: 320, top: 240 },
    });

    expect(writes).toBe(0);
    expect(transform).toBe("translate3d(10319px, 240px, 0px)");
  });

  test("measures from no offset when the host carries another transform", () => {
    const host = createHost("scale(2)");
    const editorView = mockEditorView();
    editorView.coordsAtPos = () => {
      expect(host.style.transform).toBe("");
      return { left: -9999, top: 0 };
    };

    syncImeCaretAnchor({
      hiddenHost: host,
      editorView,
      visibleCaret: { left: 320, top: 240 },
    });

    expect(host.style.transform).toBe("translate3d(10319px, 240px, 0)");
  });

  test("leaves a host that was never anchored without a transform", () => {
    const host = createHost();

    syncImeCaretAnchor({
      hiddenHost: host,
      editorView: mockEditorView({ empty: false }),
      visibleCaret: null,
    });
    resetImeCaretAnchor(host);

    expect(host.style.transform).toBe("");
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
    expect(host.style.transform).toBe("translate3d(0px, 0px, 0)");

    host.style.transform = "translate3d(10px, 20px, 0)";
    expect(
      syncImeCaretAnchor({
        hiddenHost: host,
        editorView: mockEditorView({ focused: false }),
        visibleCaret: { left: 10, top: 20 },
      }),
    ).toBe(false);
    expect(host.style.transform).toBe("translate3d(0px, 0px, 0)");
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
    expect(host.style.transform).toBe("translate3d(0px, 0px, 0)");
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
