import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { afterAll, afterEach, describe, expect, test } from "bun:test";
import type { Comment } from "@stll/folio-core/types/content";
import type { Document } from "@stll/folio-core/types/document";
import { createEmptyDocument } from "@stll/folio-core/utils/createDocument";
import { act, useLayoutEffect, useState } from "react";
import { createRoot, type Root } from "react-dom/client";

import { writeCommentAnchorIds } from "@stll/folio-core/render-dom/commentAnchorAttributes";

import { allocateCommentId, createComment } from "./commentsHelpers";
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
  controlledComments?: { current: Comment[] };
  onCommentsChange?: (comments: Comment[]) => void;
  /** Painted content root the highlight sync reads its anchors from. */
  editorContent?: HTMLElement;
  onChildLayout?: (hook: Hook) => void;
  /** Canonical committed comments, read on every render like the editor does. */
  committed?: { current: readonly Comment[] | null };
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
  controlledComments,
  onCommentsChange,
  editorContent,
  onChildLayout,
  committed,
}: MountOptions = {}): Harness => {
  let latest: Hook | null = null;
  let bump: (() => void) | null = null;
  const Child = ({ hook }: { hook: Hook }) => {
    useLayoutEffect(() => onChildLayout?.(hook), [hook]);
    return null;
  };
  const Host = () => {
    const [, setTick] = useState(0);
    const hook = useFolioComments({
      doc: doc ?? null,
      autoOpenReviewSidebar,
      anchorPositions: new Map(),
      editorContentRef: { current: editorContent ?? null },
      commentsProp: controlledComments?.current ?? commentsProp,
      onCommentsChange,
      committedComments: committed?.current,
    });
    useLayoutEffect(() => {
      bump = () => setTick((tick) => tick + 1);
      latest = hook;
    });
    return <Child hook={hook} />;
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
  test("loaded comments notify after commit once per document initialization", () => {
    const doc = createEmptyDocument();
    const loaded = [makeComment(41)];
    doc.package.document.comments = loaded;
    const changes: Comment[][] = [];
    const harness = mount({ doc, onCommentsChange: (next) => changes.push(next) });
    expect(changes).toEqual([loaded]);
    expect(harness.hook.comments).toBe(loaded);
    harness.rerender();
    expect(changes).toEqual([loaded]);
    act(() => harness.hook.resetLoadedComments());
    expect(changes).toEqual([loaded, loaded]);
  });

  test("a layout commit reapplies highlights to newly painted marks", () => {
    const editorContent = document.createElement("div");
    const harness = mount({ commentsProp: [makeComment(42)], editorContent });
    act(() => harness.hook.setActiveCommentId(42));
    const mark = document.createElement("span");
    mark.className = "layout-run-text";
    writeCommentAnchorIds(mark, [42]);
    mark.style.boxShadow = "1px 1px black";
    editorContent.append(mark);
    harness.rerender();
    expect(mark.dataset["activeComment"]).toBe("true");
    // Happy DOM does not parse the CSS-variable border shorthand. Check the
    // active marker and a supported style write instead.
    expect(mark.style.boxShadow).toBe("none");
  });

  test.each(["loaded", "controlled"] as const)(
    "reserves %s comment IDs before creating a comment",
    (source) => {
      const existingId = allocateCommentId() + 500;
      const existing = makeComment(existingId);
      const doc = createEmptyDocument();
      doc.package.document.comments = [existing];
      mount(source === "loaded" ? { doc } : { commentsProp: [existing] });
      expect(createComment("New note", "Reviewer").id).toBe(existingId + 1);
    },
  );

  test.each(["loaded", "controlled"] as const)(
    "reserves %s IDs before child layout callbacks",
    (source) => {
      const existingId = allocateCommentId() + 500;
      const existing = makeComment(existingId);
      const doc = createEmptyDocument();
      doc.package.document.comments = [existing];
      let allocated: number | undefined;
      mount({
        ...(source === "loaded" ? { doc } : { commentsProp: [existing] }),
        onChildLayout: (hook) => {
          allocated ??= hook.createComment("New note", "Reviewer").id;
        },
      });
      expect(allocated).toBe(existingId + 1);
    },
  );

  test("reserves IDs synchronously when adopting comments through the setter", () => {
    const harness = mount();
    const existingId = allocateCommentId() + 500;
    act(() => {
      harness.hook.setComments([makeComment(existingId)]);
      expect(createComment("New note", "Reviewer").id).toBe(existingId + 1);
    });
  });
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

  test("controlled: rejecting an update never changes the mirror, even before host rerender", () => {
    const initial = [makeComment(1)];
    const changes: Comment[][] = [];
    const harness = mount({
      commentsProp: initial,
      onCommentsChange: (next) => changes.push(next),
    });
    const authoritative = harness.hook.comments;
    const next = [...initial, makeComment(2)];

    act(() => {
      harness.hook.setComments(next);
      expect(harness.hook.commentsRef.current).toBe(authoritative);
      expect(harness.hook.commentsRef.current.map((comment) => comment.id)).toEqual([1]);
    });

    expect(changes).toEqual([next]);
    expect(harness.hook.isControlledComments).toBe(true);
    expect(harness.hook.comments).toBe(authoritative);

    // The host ignores the notification and rerenders its unchanged prop array.
    harness.rerender();
    expect(harness.hook.comments).toBe(authoritative);
    expect(harness.hook.commentsRef.current).toBe(authoritative);

    act(() => {
      harness.hook.setComments((previous) => {
        expect(previous).toBe(authoritative);
        return [...previous, makeComment(3)];
      });
    });
    expect(changes.map((comments) => comments.map((comment) => comment.id))).toEqual([
      [1, 2],
      [1, 3],
    ]);
  });

  test("controlled: an accepted update enters the mirror only when the host commits it", () => {
    const controlledComments = { current: [makeComment(1)] };
    const harness = mount({
      controlledComments,
      onCommentsChange: (next) => {
        controlledComments.current = next;
      },
    });
    const authoritative = harness.hook.comments;
    const next = [...authoritative, makeComment(2)];

    act(() => {
      harness.hook.setComments(next);
      expect(controlledComments.current).toBe(next);
      expect(harness.hook.commentsRef.current).toBe(authoritative);
    });

    harness.rerender();
    expect(harness.hook.comments.map((comment) => comment.id)).toEqual([1, 2]);
    expect(harness.hook.commentsRef.current).toBe(harness.hook.comments);
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

describe("useFolioComments committed canonical comments", () => {
  const hostModes = [
    { name: "uncontrolled", commentsProp: undefined },
    { name: "controlled with unchanged host props", commentsProp: [] as Comment[] },
  ] as const;

  for (const { name, commentsProp } of hostModes) {
    test(`${name}: the sidebar lists canonical creation, resolution and deletion`, () => {
      const committed: { current: readonly Comment[] | null } = { current: [] };
      const harness = mount({ commentsProp, committed });
      expect(harness.hook.visibleComments).toEqual([]);

      const created = makeComment(41);
      committed.current = [created];
      harness.rerender();
      expect(harness.hook.visibleComments).toEqual([created]);
      expect(harness.hook.commentAuthors).toEqual(["Reviewer"]);

      const resolved = { ...created, done: true };
      committed.current = [resolved];
      harness.rerender();
      expect(harness.hook.visibleComments).toEqual([resolved]);

      committed.current = [];
      harness.rerender();
      expect(harness.hook.visibleComments).toEqual([]);
      expect(harness.hook.commentAuthors).toEqual([]);
    });
  }

  for (const { name, commentsProp } of hostModes) {
    test(`${name}: loaded canonical threads open the sidebar once when enabled`, () => {
      const committed = { current: [makeComment(41)] };
      const harness = mount({ commentsProp, committed, autoOpenReviewSidebar: true });
      expect(harness.hook.showCommentsSidebar).toBe(true);
      act(() => harness.hook.setShowCommentsSidebar(false));
      committed.current = [makeComment(41), makeComment(42)];
      harness.rerender();
      expect(harness.hook.showCommentsSidebar).toBe(false);
    });

    test(`${name}: disabled auto-open and resolved canonical threads keep the sidebar closed`, () => {
      const committed = { current: [makeComment(41)] };
      expect(mount({ commentsProp, committed }).hook.showCommentsSidebar).toBe(false);
      committed.current = [{ ...makeComment(41), done: true }];
      expect(
        mount({ commentsProp, committed, autoOpenReviewSidebar: true }).hook.showCommentsSidebar,
      ).toBe(false);
    });
  }

  test("without a canonical list the host comments stay the source", () => {
    const hostComment = makeComment(7);
    const harness = mount({ commentsProp: [hostComment], committed: { current: null } });
    expect(harness.hook.visibleComments).toEqual([hostComment]);
  });
});
