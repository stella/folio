import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { afterAll, expect, mock, test } from "bun:test";
import { act, createRef, useCallback, useRef, useState } from "react";
import { IntlProvider } from "use-intl";
import { createRoot } from "react-dom/client";

import { getFolioMessages } from "@stll/folio-core/i18n/messages";
import { DEFAULT_OUTLINE_DEPTH, filterHeadingsByDepth } from "@stll/folio-core/utils/outlineDepth";
import type { HeadingInfo } from "@stll/folio-core/utils/headingCollector";
import { DocumentOutline } from "./DocumentOutline";

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

test("defaults to two outline levels and reports depth selector changes", async () => {
  const onOutlineDepthChange = mock((_depth: 2 | 3 | "all") => {});
  const OutlineHarness = () => {
    const [depth, setDepth] = useState(DEFAULT_OUTLINE_DEPTH);
    const scrollContainerRef = useRef<HTMLDivElement>(null);
    const onDepthChange = useCallback(
      (nextDepth: 2 | 3 | "all") => {
        onOutlineDepthChange(nextDepth);
        setDepth(nextDepth);
      },
      [onOutlineDepthChange],
    );
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
