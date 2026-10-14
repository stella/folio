import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { afterAll, expect, test } from "bun:test";
import type { Comment } from "@stll/folio-core/types/content";
import {
  commentAnchorSelector,
  writeCommentAnchorIds,
} from "@stll/folio-core/render-dom/commentAnchorAttributes";
import { act, createRef } from "react";
import { createRoot } from "react-dom/client";
import { IntlProvider } from "use-intl";

import { getFolioMessages } from "@stll/folio-core/i18n/messages";

import { CommentsSidebar } from "./CommentsSidebar";

const previousActEnvironment = Reflect.get(globalThis, "IS_REACT_ACT_ENVIRONMENT");
Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", true);
afterAll(() => {
  Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", previousActEnvironment);
  GlobalRegistrator.unregister();
});

// Existing panel geometry tests did not vary DOM and layout anchor coordinates.
const comment = {
  id: 1,
  author: "Reviewer",
  date: "2026-01-01T00:00:00Z",
  content: [],
} satisfies Comment;
const comments = [comment];

const setRect = (element: HTMLElement, top: number) => {
  element.getBoundingClientRect = () => new DOMRect(0, top, 800, 100);
};

type RenderDrawerOptions = {
  scrollElement: HTMLDivElement;
  anchorPositions: Map<string, number>;
  activeCommentId?: number;
  surface?: "column" | "drawer";
  threadComments?: Comment[];
};

const renderDrawer = async ({
  scrollElement,
  anchorPositions,
  activeCommentId,
  surface = "drawer",
  threadComments = comments,
}: RenderDrawerOptions) => {
  scrollElement.setAttribute("data-folio-scroll", "");
  const host = document.createElement("div");
  document.body.append(host);
  const root = createRoot(host);
  const editorContainerRef = createRef<HTMLDivElement>();
  editorContainerRef.current = scrollElement;
  let selectedCommentId = activeCommentId ?? null;
  // oxlint-disable-next-line react-perf/jsx-no-new-function-as-prop -- created once per mounted test root and reused across every render
  const activateComment = (commentId: number | null) => {
    selectedCommentId = commentId;
    render();
  };
  const render = () => {
    root.render(
      <IntlProvider locale="en" timeZone="UTC" messages={getFolioMessages("en")}>
        <CommentsSidebar
          comments={threadComments}
          editorContainerRef={editorContainerRef}
          anchorPositions={anchorPositions}
          activeCommentId={selectedCommentId}
          onCommentClick={activateComment}
          surface={surface}
        />
      </IntlProvider>,
    );
  };
  await act(async () => render());
  return { host, root, activateComment };
};

const measurePositions = async () => {
  await act(async () => {
    await new Promise((resolve) => setTimeout(resolve, 60));
  });
};

test("reactivating a thread after an outside click removes the reply preview clamp", async () => {
  const replyText = "Preview line one\nPreview line two\nSearch match beyond the preview";
  const reply = {
    id: 2,
    parentId: comment.id,
    author: "Reviewer",
    date: "2026-01-01T00:00:00Z",
    content: [
      {
        type: "paragraph",
        formatting: {},
        content: [
          {
            type: "run",
            formatting: {},
            content: [{ type: "text", text: replyText }],
          },
        ],
      },
    ],
  } satisfies Comment;
  const scrollElement = document.createElement("div");
  const pages = document.createElement("div");
  pages.className = "paged-editor__pages";
  scrollElement.append(pages);
  const { host, root, activateComment } = await renderDrawer({
    scrollElement,
    anchorPositions: new Map(),
    activeCommentId: comment.id,
    threadComments: [comment, reply],
  });
  try {
    const replyBody = Array.from(host.querySelectorAll<HTMLDivElement>("div")).find(
      (element) => element.textContent === replyText && element.style.lineHeight === "17px",
    );
    expect(replyBody).toBeDefined();
    expect(replyBody?.style.overflow).not.toBe("hidden");

    // Previous fixtures never collapsed an active thread outside the sidebar.
    await act(async () => pages.click());
    expect(replyBody?.style.overflow).toBe("hidden");

    await act(async () => activateComment(comment.id));
    expect(replyBody?.style.overflow).not.toBe("hidden");
    expect(host.querySelector(".docx-comment-card input[readonly]")).not.toBeNull();
  } finally {
    await act(async () => root.unmount());
    host.remove();
    scrollElement.remove();
  }
});

test("drawer expansion prefers the rendered comment anchor over layout positions", async () => {
  const scrollElement = document.createElement("div");
  Object.defineProperty(scrollElement, "clientHeight", { configurable: true, value: 300 });
  scrollElement.scrollTop = 300;
  setRect(scrollElement, 100);
  const pages = document.createElement("div");
  pages.className = "paged-editor__pages";
  const anchor = document.createElement("span");
  writeCommentAnchorIds(anchor, [comment.id]);
  setRect(anchor, 700);
  pages.append(anchor);
  scrollElement.append(pages);
  expect(pages.querySelector(commentAnchorSelector(comment.id))).toBe(anchor);

  const { host, root } = await renderDrawer({
    scrollElement,
    anchorPositions: new Map([["comment-1", 120]]),
  });
  try {
    await measurePositions();
    await act(async () => {
      host.querySelector<HTMLElement>(".docx-comment-card")?.click();
    });
    expect(scrollElement.scrollTop).toBe(800);
  } finally {
    await act(async () => root.unmount());
    host.remove();
    scrollElement.remove();
  }
});

test("first positioned cards fade in before subsequent geometry moves animate", async () => {
  const scrollElement = document.createElement("div");
  const pages = document.createElement("div");
  pages.className = "paged-editor__pages";
  scrollElement.append(pages);
  const anchorPositions = new Map([["comment-1", 120]]);
  const { host, root } = await renderDrawer({ scrollElement, anchorPositions, surface: "column" });
  try {
    await measurePositions();
    const card = host.querySelector<HTMLElement>(".docx-comment-card");
    expect(card?.style.top).toBe("120px");
    expect(card?.style.transition).not.toContain("top");

    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 410));
    });
    anchorPositions.set("comment-1", 240);
    await act(async () => scrollElement.dispatchEvent(new Event("scroll")));
    await measurePositions();
    expect(card?.style.top).toBe("240px");
    expect(card?.style.transition).toContain("top 0.15s ease");
  } finally {
    await act(async () => root.unmount());
    host.remove();
    scrollElement.remove();
  }
});

test("drawer expansion uses layout positions when the comment anchor is not rendered", async () => {
  const scrollElement = document.createElement("div");
  Object.defineProperty(scrollElement, "clientHeight", { configurable: true, value: 300 });
  setRect(scrollElement, 100);
  const pages = document.createElement("div");
  pages.className = "paged-editor__pages";
  scrollElement.append(pages);
  expect(pages.querySelector(commentAnchorSelector(comment.id))).toBeNull();

  const { host, root } = await renderDrawer({
    scrollElement,
    anchorPositions: new Map([["comment-1", 420]]),
  });
  try {
    await measurePositions();
    await act(async () => {
      host.querySelector<HTMLElement>(".docx-comment-card")?.click();
    });
    expect(scrollElement.scrollTop).toBe(320);
  } finally {
    await act(async () => root.unmount());
    host.remove();
    scrollElement.remove();
  }
});

test("mounting with an active comment expands its reply input and permits collapsing it", async () => {
  const scrollElement = document.createElement("div");
  const { host, root } = await renderDrawer({
    scrollElement,
    anchorPositions: new Map(),
    activeCommentId: comment.id,
  });
  try {
    // Earlier fixtures only mounted inactive cards and then clicked to expand.
    const card = host.querySelector<HTMLElement>(".docx-comment-card");
    expect(card).not.toBeNull();
    expect(card?.querySelector("input[readonly]")).not.toBeNull();

    await act(async () => card?.click());
    expect(card?.querySelector("input[readonly]")).toBeNull();

    await act(async () => card?.click());
    expect(card?.querySelector("input[readonly]")).not.toBeNull();
  } finally {
    await act(async () => root.unmount());
    host.remove();
    scrollElement.remove();
  }
});
