import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { afterAll, afterEach, describe, expect, test } from "bun:test";
import type { Comment } from "@stll/folio-core/types/content";
import type { Document } from "@stll/folio-core/types/document";
import { createEmptyDocument } from "@stll/folio-core/utils/createDocument";
import { act, useState } from "react";
import { createRoot, type Root } from "react-dom/client";

import { writeCommentAnchorIds } from "@stll/folio-core/render-dom/commentAnchorAttributes";

import { useFolioComments } from "./useFolioComments";

// React only silences its "not wrapped in act" warning when this flag is set.
(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

// Unmount every root so the hook's effect cleanup cancels its pending frame
// and timer; otherwise that work outlives the test that scheduled it.
const roots: Root[] = [];

afterEach(() => {
  act(() => {
    for (const root of roots.splice(0)) {
      root.unmount();
    }
  });
});

afterAll(() => {
  GlobalRegistrator.unregister();
});

type Hook = ReturnType<typeof useFolioComments>;

type Harness = {
  readonly hook: Hook;
  /** Rerender the host without touching any comment input. */
  rerender: () => void;
};

type MountOptions = {
  doc?: Document;
  autoOpenReviewSidebar?: boolean;
  commentsProp?: Comment[];
  onCommentsChange?: (comments: Comment[]) => void;
  /** Painted content root the highlight sync reads its anchors from. */
  editorContent?: HTMLElement;
};

const makeComment = (id: number): Comment => ({
  id,
  author: "Reviewer",
  date: "2026-01-01T00:00:00Z",
  content: [],
});

const mount = ({
  doc,
  autoOpenReviewSidebar = false,
  commentsProp,
  onCommentsChange,
  editorContent,
}: MountOptions = {}): Harness => {
  let latest: Hook | null = null;
  let bump: (() => void) | null = null;
  const Host = () => {
    const [, setTick] = useState(0);
    bump = () => setTick((tick) => tick + 1);
    latest = useFolioComments({
      doc: doc ?? null,
      autoOpenReviewSidebar,
      anchorPositions: new Map(),
      editorContentRef: { current: editorContent ?? null },
      commentsProp,
      onCommentsChange,
    });
    return null;
  };
  const root = createRoot(document.createElement("div"));
  roots.push(root);
  act(() => root.render(<Host />));
  return {
    get hook() {
      if (!latest) {
        throw new Error("hook not mounted");
      }
      return latest;
    },
    rerender: () => {
      act(() => bump?.());
    },
  };
};

describe("useFolioComments.setComments", () => {
  test("uncontrolled: a mutation survives an unrelated host rerender", () => {
    const changes: Comment[][] = [];
    const harness = mount({ onCommentsChange: (next) => changes.push(next) });
    const next = [makeComment(1)];

    act(() => harness.hook.setComments(next));

    expect(harness.hook.comments).toBe(next);
    expect(harness.hook.commentsRef.current).toBe(next);
    expect(changes).toEqual([next]);

    harness.rerender();

    expect(harness.hook.comments).toBe(next);
    expect(harness.hook.commentsRef.current).toBe(next);
  });

  test("uncontrolled: the ref is updated synchronously for same-tick readers", () => {
    const harness = mount();
    const next = [makeComment(1)];

    act(() => {
      harness.hook.setComments(next);
      expect(harness.hook.commentsRef.current).toBe(next);
    });
  });

  test("functional updates read the latest value, including within one tick", () => {
    const harness = mount();

    act(() => {
      harness.hook.setComments((previous) => [...previous, makeComment(1)]);
      harness.hook.setComments((previous) => [...previous, makeComment(2)]);
    });

    expect(harness.hook.comments.map((comment) => comment.id)).toEqual([1, 2]);
  });

  test("the identical reference is a no-op and does not notify", () => {
    const changes: Comment[][] = [];
    const harness = mount({ onCommentsChange: (next) => changes.push(next) });
    const next = [makeComment(1)];

    act(() => harness.hook.setComments(next));
    act(() => harness.hook.setComments(next));
    act(() => harness.hook.setComments(() => next));

    expect(changes).toEqual([next]);
  });

  test("controlled: notifies the host and leaves the prop authoritative", () => {
    const initial = [makeComment(1)];
    const changes: Comment[][] = [];
    const harness = mount({
      commentsProp: initial,
      onCommentsChange: (next) => changes.push(next),
    });
    const next = [...initial, makeComment(2)];

    act(() => harness.hook.setComments(next));

    expect(changes).toEqual([next]);
    expect(harness.hook.isControlledComments).toBe(true);

    // The host did not apply the change, so the next render reads the prop.
    harness.rerender();
    expect(harness.hook.comments.map((comment) => comment.id)).toEqual([1]);
  });
});

describe("useFolioComments highlight sync", () => {
  /** A painted content root: one run in comment 7, one in both 7 and 9. */
  const paintedRoot = (): { root: HTMLElement; inOne: HTMLElement; inBoth: HTMLElement } => {
    const root = document.createElement("div");
    const inOne = document.createElement("span");
    const inBoth = document.createElement("span");
    for (const [element, commentIds] of [
      [inOne, [7]],
      [inBoth, [7, 9]],
    ] as const) {
      element.className = "layout-run-text";
      writeCommentAnchorIds(element, commentIds);
      root.append(element);
    }
    return { root, inOne, inBoth };
  };

  test("a run inside overlapping ranges goes active for either comment", () => {
    const { root, inOne, inBoth } = paintedRoot();
    const harness = mount({
      commentsProp: [makeComment(7), makeComment(9)],
      editorContent: root,
    });

    act(() => harness.hook.setActiveCommentId(7));
    expect(inBoth.dataset["activeComment"]).toBe("true");
    expect(inOne.dataset["activeComment"]).toBe("true");

    // The inner comment is the one the user is on now; the run is in it too.
    act(() => harness.hook.setActiveCommentId(9));
    expect(inBoth.dataset["activeComment"]).toBe("true");
    expect(inOne.dataset["activeComment"]).toBeUndefined();
  });
});

describe("useFolioComments auto-open", () => {
  const openOnLoad = (comments: Comment[]) => {
    const doc = createEmptyDocument();
    doc.package.document.comments = comments;
    return mount({ doc, autoOpenReviewSidebar: true }).hook.showCommentsSidebar;
  };

  test("opens the sidebar when a thread would show a card", () => {
    expect(openOnLoad([makeComment(1)])).toBe(true);
  });

  test("keeps the sidebar closed when it would be empty", () => {
    // Resolved threads and replies have no card of their own, so the panel
    // would only say "No comments yet."
    const resolved = { ...makeComment(1), done: true };
    const reply = { ...makeComment(2), parentId: 1 };
    expect(openOnLoad([resolved, reply])).toBe(false);
  });
});
