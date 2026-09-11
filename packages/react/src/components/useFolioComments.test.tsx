import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { afterAll, afterEach, describe, expect, test } from "bun:test";
import type { Comment } from "@stll/folio-core/types/content";
import { act, useState } from "react";
import { createRoot, type Root } from "react-dom/client";

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
  commentsProp?: Comment[];
  onCommentsChange?: (comments: Comment[]) => void;
};

const makeComment = (id: number): Comment => ({
  id,
  author: "Reviewer",
  date: "2026-01-01T00:00:00Z",
  content: [],
});

const mount = ({ commentsProp, onCommentsChange }: MountOptions = {}): Harness => {
  let latest: Hook | null = null;
  let bump: (() => void) | null = null;
  const Host = () => {
    const [, setTick] = useState(0);
    bump = () => setTick((tick) => tick + 1);
    latest = useFolioComments({
      doc: null,
      autoOpenReviewSidebar: false,
      anchorPositions: new Map(),
      editorContentRef: { current: null },
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
