import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { afterAll, expect, mock, test } from "bun:test";
import { act, createRef, useCallback, useRef, useState } from "react";
import { IntlProvider } from "use-intl";
import { createRoot } from "react-dom/client";

import { getFolioMessages } from "@stll/folio-core/i18n/messages";
import { DEFAULT_OUTLINE_DEPTH, filterHeadingsByDepth } from "@stll/folio-core/utils/outlineDepth";
import type { HeadingInfo } from "@stll/folio-core/utils/headingCollector";
import type { Paragraph } from "@stll/folio-core/types/document";
import { createEmptyDocument } from "@stll/folio-core/utils/createDocument";
import type { OutlineDepth } from "@stll/folio-core/utils/outlineDepth";
import { DocumentOutline, type DocumentOutlineSurface } from "./DocumentOutline";
import { DocxEditor } from "./DocxEditor";
import type { DocxEditorRef } from "./DocxEditor.props";

const previousActEnvironment = Reflect.get(globalThis, "IS_REACT_ACT_ENVIRONMENT");
Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", true);
afterAll(() => {
  Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", previousActEnvironment);
  GlobalRegistrator.unregister();
});

const headings: HeadingInfo[] = [
  { text: "First level", level: 0, pmPos: 1 },
  { text: "Second level", level: 1, pmPos: 2 },
  { text: "Third level", level: 2, pmPos: 3 },
  { text: "Deep level", level: 8, pmPos: 4 },
];
const EMPTY_HEADINGS: HeadingInfo[] = [];
const onJumpNoop = () => {};
const onOutlineDepthChangeNoop = (_depth: OutlineDepth) => {};
const outlineParagraph = (styleId: string, text: string): Paragraph => ({
  type: "paragraph",
  formatting: { styleId },
  content: [{ type: "run", content: [{ type: "text", text }] }],
});
const createOutlineDocument = () => {
  const document = createEmptyDocument();
  document.package.document.content = [
    outlineParagraph("Heading1", "First level"),
    outlineParagraph("Heading2", "Second level"),
    outlineParagraph("Heading3", "Third level"),
  ];
  return document;
};

test("defaults to two outline levels and reports depth selector changes", async () => {
  const onOutlineDepthChange = mock((_depth: 2 | 3 | "all") => {});
  const OutlineHarness = () => {
    const [depth, setDepth] = useState<OutlineDepth>(DEFAULT_OUTLINE_DEPTH);
    const scrollContainerRef = useRef<HTMLDivElement>(null);
    const onDepthChange = useCallback((nextDepth: 2 | 3 | "all") => {
      onOutlineDepthChange(nextDepth);
      setDepth(nextDepth);
    }, []);
    return (
      <IntlProvider locale="en" timeZone="UTC" messages={getFolioMessages("en")}>
        <DocumentOutline
          activeId={null}
          available
          docSize={10}
          headings={filterHeadingsByDepth(headings, depth)}
          onOutlineDepthChange={onDepthChange}
          onJump={onJumpNoop}
          outlineDepth={depth}
          scrollContainerRef={scrollContainerRef}
          surface="expanded"
        />
      </IntlProvider>
    );
  };
  const container = document.createElement("div");
  document.body.append(container);
  const root = createRoot(container);

  try {
    await act(async () => root.render(<OutlineHarness />));

    const itemLabels = () =>
      [...container.querySelectorAll(".folio-outline-item-label")].map((node) => node.textContent);
    const select = container.querySelector("select");
    expect(select?.value).toBe("2");
    expect(itemLabels()).toEqual(["First level", "Second level"]);
    if (!select) throw new Error("outline depth selector missing");
    await act(async () => {
      select.value = "3";
      select.dispatchEvent(new Event("change", { bubbles: true }));
    });
    expect(onOutlineDepthChange).toHaveBeenCalledWith(3);
    expect(itemLabels()).toContain("Third level");
    expect(itemLabels()).not.toContain("Deep level");

    await act(async () => {
      select.value = "all";
      select.dispatchEvent(new Event("change", { bubbles: true }));
    });
    expect(onOutlineDepthChange).toHaveBeenCalledWith("all");
    expect(itemLabels()).toContain("Deep level");

    await act(async () =>
      root.render(
        <IntlProvider locale="en" timeZone="UTC" messages={getFolioMessages("en")}>
          <DocumentOutline
            activeId={null}
            available
            docSize={10}
            headings={EMPTY_HEADINGS}
            onOutlineDepthChange={onOutlineDepthChange}
            onJump={onJumpNoop}
            outlineDepth={DEFAULT_OUTLINE_DEPTH}
            scrollContainerRef={createRef<HTMLDivElement>()}
            surface="expanded"
          />
        </IntlProvider>,
      ),
    );
    expect(container.querySelector("select")).not.toBeNull();
  } finally {
    await act(async () => root.unmount());
    container.remove();
  }
});

test("DocxEditor defaults to two levels and forwards depth changes to the host", async () => {
  const outlineDocument = createOutlineDocument();
  const onOutlineDepthChange = mock((_depth: OutlineDepth) => {});
  const editorRef = createRef<DocxEditorRef>();
  const container = document.createElement("div");
  document.body.append(container);
  const root = createRoot(container);
  const previousClientWidth = Object.getOwnPropertyDescriptor(HTMLElement.prototype, "clientWidth");

  try {
    Object.defineProperty(HTMLElement.prototype, "clientWidth", {
      configurable: true,
      get(this: HTMLElement) {
        if (this.classList.contains("folio-panels-row")) return 1440;
        return previousClientWidth?.get?.call(this) ?? 0;
      },
    });
    await act(async () =>
      root.render(
        <IntlProvider locale="en" timeZone="UTC" messages={getFolioMessages("en")}>
          <DocxEditor
            ref={editorRef}
            document={outlineDocument}
            onOutlineDepthChange={onOutlineDepthChange}
            showToolbar={false}
          />
        </IntlProvider>,
      ),
    );
    await act(async () => editorRef.current?.ensureEditorView({ focus: false }));

    const itemLabels = () =>
      [...container.querySelectorAll(".folio-outline-item-label")].map((node) => node.textContent);
    expect(itemLabels()).toContain("First level");
    expect(itemLabels()).toContain("Second level");
    expect(itemLabels()).not.toContain("Third level");

    const select = container.querySelector(".folio-outline select");
    if (!(select instanceof HTMLSelectElement)) throw new Error("outline depth selector missing");
    await act(async () => {
      select.value = "all";
      select.dispatchEvent(new Event("change", { bubbles: true }));
    });
    expect(onOutlineDepthChange).toHaveBeenCalledWith("all");
    expect(itemLabels()).toContain("Third level");
  } finally {
    await act(async () => root.unmount());
    container.remove();
    if (previousClientWidth) {
      Object.defineProperty(HTMLElement.prototype, "clientWidth", previousClientWidth);
    } else {
      Reflect.deleteProperty(HTMLElement.prototype, "clientWidth");
    }
  }
});

test("Escape from expanded outline restores focus to the newly mounted rail toggle", async () => {
  const container = document.createElement("div");
  document.body.append(container);
  const root = createRoot(container);
  const OutlineFocusHarness = () => {
    const [surface, setSurface] = useState<DocumentOutlineSurface>("rail");
    const scrollContainerRef = useRef<HTMLDivElement>(null);
    const expand = useCallback(() => setSurface("expanded"), []);
    const close = useCallback(() => setSurface("rail"), []);
    return (
      <IntlProvider locale="en" timeZone="UTC" messages={getFolioMessages("en")}>
        <DocumentOutline
          activeId={null}
          available
          docSize={10}
          expanded={surface === "expanded"}
          headings={EMPTY_HEADINGS}
          onClose={close}
          onExpand={expand}
          onJump={onJumpNoop}
          onOutlineDepthChange={onOutlineDepthChangeNoop}
          outlineDepth={DEFAULT_OUTLINE_DEPTH}
          scrollContainerRef={scrollContainerRef}
          surface={surface}
        />
      </IntlProvider>
    );
  };

  try {
    await act(async () => root.render(<OutlineFocusHarness />));
    const opener = container.querySelector<HTMLButtonElement>(
      "[data-testid='folio-outline-expand']",
    );
    if (!opener) throw new Error("outline rail toggle missing");
    await act(async () => {
      opener.focus();
      opener.click();
    });
    expect(
      document.activeElement?.closest('[data-folio-outline-surface="expanded"]'),
    ).not.toBeNull();

    await act(async () => {
      document.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true }));
    });
    const restoredToggle = container.querySelector<HTMLButtonElement>(
      "[data-testid='folio-outline-expand']",
    );
    expect(restoredToggle).not.toBe(opener);
    expect(restoredToggle?.isConnected).toBe(true);
    expect(document.activeElement).toBe(restoredToggle);
  } finally {
    await act(async () => root.unmount());
    container.remove();
  }
});
