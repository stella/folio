import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { act, Profiler, useCallback, useRef, useState } from "react";
import { createRoot } from "react-dom/client";

import { panic } from "better-result";
import { afterAll, expect, spyOn, test } from "bun:test";

import { createFolioEditorEmitter } from "@stll/folio-core/controller/folioEditorEvents";
import type { Layout } from "@stll/folio-core/layout-engine/types";
import type { Document } from "@stll/folio-core/types/document";
import { createEmptyDocument } from "@stll/folio-core/utils/createDocument";
import {
  createEmptyHeaderFooter,
  resolveHeaderFooterContent,
} from "@stll/folio-core/utils/headerFooter";

import { useHistory } from "../../hooks/useHistory";
import type { PagedEditorRef } from "../../paged-editor/PagedEditor";
import { useHeaderFooterEditor } from "./useHeaderFooterEditor";
import { useZoomAndPageInfo } from "./useZoomAndPageInfo";

const previousActEnvironment = Reflect.get(
  globalThis,
  "IS_REACT_ACT_ENVIRONMENT",
);
Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", true);
afterAll(() => {
  Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", previousActEnvironment);
  GlobalRegistrator.unregister();
});
const returnDocument = (document: Document) => document;
const noView = () => null;

const HeaderHarness = ({ document }: { document: Document | null }) => {
  const history = useHistory(document, { enableKeyboardShortcuts: false });
  const editor = useHeaderFooterEditor({
    history: { ...history, state: document },
    pushDocument: returnDocument,
    getHfView: noView,
  });
  return (
    <output>
      {JSON.stringify({
        header: editor.activeHeaderRId,
        footer: editor.activeFooterRId,
        headerContent: editor.headerContent,
        footerContent: editor.footerContent,
      })}
    </output>
  );
};

test("header/footer content and active ids survive transient empty documents and follow the next document", async () => {
  const first =
    createEmptyHeaderFooter(createEmptyDocument(), "header", false) ??
    panic("Header fixture missing");
  const second =
    createEmptyHeaderFooter(createEmptyDocument(), "footer", false) ??
    panic("Footer fixture missing");
  const container = document.createElement("div");
  const root = createRoot(container);
  try {
    await act(async () => root.render(<HeaderHarness document={first} />));
    const firstOutput = container.textContent;
    expect(firstOutput).toContain(
      resolveHeaderFooterContent(first.package).activeHeaderRId ??
        panic("Header id missing"),
    );
    for (let index = 0; index < 3; index++) {
      await act(async () => root.render(<HeaderHarness document={null} />));
      expect(container.textContent).toBe(firstOutput);
    }
    await act(async () => root.render(<HeaderHarness document={second} />));
    const secondOutput = container.textContent;
    expect(secondOutput).toContain(
      resolveHeaderFooterContent(second.package).activeFooterRId ??
        panic("Footer id missing"),
    );
    expect(secondOutput).not.toBe(firstOutput);
    await act(async () => root.render(<HeaderHarness document={null} />));
    expect(container.textContent).toBe(secondOutput);
  } finally {
    await act(async () => root.unmount());
  }
});

const createCommitProbe = () => {
  const commits: string[] = [];
  return {
    commits,
    onRender: (_id: string, phase: string) => commits.push(phase),
  };
};

const ZoomHarness = ({ mounted }: { mounted: boolean }) => {
  const scrollContainerRef = useRef<HTMLDivElement>(null);
  const pagedEditorRef = useRef<PagedEditorRef>(null);
  const [scrollContainer, setScrollContainer] = useState<HTMLDivElement | null>(
    null,
  );
  const attach = useCallback((element: HTMLDivElement | null) => {
    scrollContainerRef.current = element;
    setScrollContainer(element);
  }, []);
  const { scrollPageInfo } = useZoomAndPageInfo({
    scrollContainerRef,
    scrollContainer,
    pagedEditorRef,
    editorController: null,
    initialZoom: 1,
  });
  return (
    <>
      {mounted && <div ref={attach} data-testid="scroll-container" />}
      <output>{String(scrollPageInfo.visible)}</output>
    </>
  );
};

test("the committed scroll container attaches one listener without per-scroll renders before layout", async () => {
  const container = document.createElement("div");
  const root = createRoot(container);
  const add = spyOn(HTMLDivElement.prototype, "addEventListener");
  const remove = spyOn(HTMLDivElement.prototype, "removeEventListener");
  const { commits, onRender } = createCommitProbe();
  const render = (mounted: boolean) => (
    <Profiler id="zoom" onRender={onRender}>
      <ZoomHarness mounted={mounted} />
    </Profiler>
  );
  try {
    await act(async () => root.render(render(false)));
    await act(async () => root.render(render(true)));
    const scrollContainer =
      container.querySelector<HTMLDivElement>(
        '[data-testid="scroll-container"]',
      ) ?? panic("Scroll container missing");
    const scrollCalls = add.mock.calls.filter(([type]) => type === "scroll");
    expect(scrollCalls.length).toBe(1);
    const commitCount = commits.length;
    await act(async () => {
      for (let index = 0; index < 12; index++)
        scrollContainer.dispatchEvent(new Event("scroll"));
    });
    expect(commits.length).toBe(commitCount);
    expect(add.mock.calls.filter(([type]) => type === "scroll").length).toBe(1);
    await act(async () => root.render(render(false)));
    expect(remove.mock.calls.filter(([type]) => type === "scroll").length).toBe(
      1,
    );
  } finally {
    await act(async () => root.unmount());
    add.mockRestore();
    remove.mockRestore();
  }
});

type InitialFitHarnessProps = {
  width: number;
  editorController: NonNullable<
    Parameters<typeof useZoomAndPageInfo>[0]["editorController"]
  >;
};

const InitialFitHarness = ({
  width,
  editorController,
}: InitialFitHarnessProps) => {
  const scrollContainerRef = useRef<HTMLDivElement>(null);
  const pagedEditorRef = useRef<PagedEditorRef>(null);
  const [scrollContainer, setScrollContainer] = useState<HTMLDivElement | null>(
    null,
  );
  const attach = useCallback(
    (element: HTMLDivElement | null) => {
      if (element) {
        Object.defineProperties(element, {
          clientWidth: { value: width, configurable: true },
          offsetWidth: { value: width, configurable: true },
        });
      }
      scrollContainerRef.current = element;
      setScrollContainer(element);
    },
    [width],
  );
  const { zoom, isInitialFitPending } = useZoomAndPageInfo({
    scrollContainerRef,
    scrollContainer,
    pagedEditorRef,
    editorController,
    initialZoom: "fit-width",
  });

  return <div ref={attach}>{isInitialFitPending ? "loading" : zoom}</div>;
};

const initialFitController = () => {
  const events = createFolioEditorEmitter();
  let currentLayout: Layout | null = null;
  return {
    getLayout: () => currentLayout,
    on: events.on,
    publishLayout: (layout: Layout) => {
      currentLayout = layout;
      events.emit("layoutComplete", layout);
    },
  };
};

test.each([
  { width: 656, expectedZoom: 0.8 },
  { width: 816, expectedZoom: 1 },
])(
  "a fit-width editor waits for the first fitted page at zoom $expectedZoom",
  async ({ width, expectedZoom }) => {
    const editorController = initialFitController();
    const pageSize = { w: 800, h: 1000 };
    const layout: Layout = {
      pageSize,
      pages: [
        {
          number: 1,
          logicalNumber: 1,
          size: pageSize,
          margins: { top: 80, right: 80, bottom: 80, left: 80 },
          fragments: [],
        },
      ],
    };
    const container = document.createElement("div");
    const root = createRoot(container);

    try {
      await act(async () => {
        root.render(
          <InitialFitHarness
            width={width}
            editorController={editorController}
          />,
        );
      });
      expect(container.textContent).toBe("loading");

      await act(async () => {
        editorController.publishLayout(layout);
      });
      expect(container.textContent).toBe(String(expectedZoom));
    } finally {
      await act(async () => root.unmount());
    }
  },
);
