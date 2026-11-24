import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { act, createRef } from "react";
import { createRoot } from "react-dom/client";

import { panic } from "better-result";
import { afterAll, expect, test } from "bun:test";
import { IntlProvider } from "use-intl";

import { getFolioMessages } from "@stll/folio-core/i18n/messages";
import { createEmptyDocument } from "@stll/folio-core/utils/createDocument";

import { PagedEditor } from "../paged-editor/PagedEditor";
import type { PagedEditorRef } from "../paged-editor/PagedEditor";
import { DocxEditor } from "./DocxEditor";

const previousActEnvironment = Reflect.get(
  globalThis,
  "IS_REACT_ACT_ENVIRONMENT",
);
Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", true);
afterAll(() => {
  Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", previousActEnvironment);
  GlobalRegistrator.unregister();
});

test("replacing the document resets the scroll position", async () => {
  const container = document.createElement("div");
  document.body.append(container);
  const root = createRoot(container);
  const renderDocument = async (text: string) => {
    await act(async () => {
      root.render(
        <IntlProvider
          locale="en"
          timeZone="UTC"
          messages={getFolioMessages("en")}
        >
          <DocxEditor
            accessibleName="Document editor"
            document={createEmptyDocument({ initialText: text })}
            showToolbar={false}
          />
        </IntlProvider>,
      );
    });
  };
  try {
    await renderDocument("First document");
    const scrollContainer = container.querySelector("[data-folio-scroll]");
    if (!(scrollContainer instanceof HTMLElement)) {
      panic("The editor did not render its scroll container");
    }
    scrollContainer.scrollTop = 4321;
    scrollContainer.scrollLeft = 17;

    await renderDocument("Replacement document");

    expect(scrollContainer.scrollTop).toBe(0);
    expect(scrollContainer.scrollLeft).toBe(0);
  } finally {
    await act(async () => root.unmount());
    container.remove();
  }
});

const documentIO = {
  getDocx: async () => null,
  loadDocument: () => {},
  loadDocx: async () => {},
};

for (const existingMarker of [undefined, "host-owned"]) {
  test(`external paged-editor scroll roots preserve marker ownership (${existingMarker ?? "unmarked"})`, async () => {
    const scrollRoot = document.createElement("div");
    if (existingMarker !== undefined)
      scrollRoot.setAttribute("data-folio-scroll", existingMarker);
    const container = document.createElement("div");
    scrollRoot.append(container);
    document.body.append(scrollRoot);
    const root = createRoot(container);
    const editor = createRef<PagedEditorRef>();
    const scrollContainerRef = createRef<HTMLDivElement>();
    scrollContainerRef.current = scrollRoot;
    try {
      await act(async () => {
        root.render(
          <IntlProvider
            locale="en"
            timeZone="UTC"
            messages={getFolioMessages("en")}
          >
            <PagedEditor
              accessibleName="Document editor"
              ref={editor}
              document={null}
              documentIdentity="scroll-root-lifecycle"
              documentIO={documentIO}
              markupView="all-markup"
              scrollContainerRef={scrollContainerRef}
            />
          </IntlProvider>,
        );
      });
      expect(editor.current?.getScrollRoot()).toBe(scrollRoot);
      expect(scrollRoot.getAttribute("data-folio-scroll")).toBe(
        existingMarker ?? "",
      );
      await act(async () => root.unmount());
      expect(scrollRoot.getAttribute("data-folio-scroll")).toBe(
        existingMarker ?? null,
      );
    } finally {
      scrollRoot.remove();
    }
  });
}
