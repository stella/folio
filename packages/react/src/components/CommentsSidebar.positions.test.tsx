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

const renderDrawer = async (
  scrollElement: HTMLDivElement,
  anchorPositions: Map<string, number>,
) => {
  const host = document.createElement("div");
  document.body.append(host);
  const root = createRoot(host);
  const editorContainerRef = createRef<HTMLDivElement>();
  editorContainerRef.current = scrollElement;
  await act(async () => {
    root.render(
      <IntlProvider locale="en" timeZone="UTC" messages={getFolioMessages("en")}>
        <CommentsSidebar
          comments={comments}
          editorContainerRef={editorContainerRef}
          anchorPositions={anchorPositions}
          surface="drawer"
        />
      </IntlProvider>,
    );
  });
  return { host, root };
};

const measurePositions = async () => {
  await act(async () => {
    await new Promise((resolve) => setTimeout(resolve, 60));
  });
};

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

  const { host, root } = await renderDrawer(scrollElement, new Map([["comment-1", 120]]));
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

test("drawer expansion uses layout positions when the comment anchor is not rendered", async () => {
  const scrollElement = document.createElement("div");
  Object.defineProperty(scrollElement, "clientHeight", { configurable: true, value: 300 });
  setRect(scrollElement, 100);
  const pages = document.createElement("div");
  pages.className = "paged-editor__pages";
  scrollElement.append(pages);
  expect(pages.querySelector(commentAnchorSelector(comment.id))).toBeNull();

  const { host, root } = await renderDrawer(scrollElement, new Map([["comment-1", 420]]));
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
