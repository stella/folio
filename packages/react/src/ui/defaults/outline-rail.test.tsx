import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { afterAll, expect, test } from "bun:test";
import { act, createRef, Profiler } from "react";
import { createRoot } from "react-dom/client";

import { createBuiltInStyleIndex } from "@stll/folio-core/docx/builtInStyles";
import {
  HEADING_COLLECTOR_DOCUMENT,
  HEADING_COLLECTOR_STYLES,
} from "@stll/folio-core/utils/fixtures/headingCollector.synthetic";
import { collectHeadings } from "@stll/folio-core/utils/headingCollector";

import type { FolioOutlineRailProps, OutlineItem } from "../folio-ui";
import { DefaultOutlineRail } from "./outline-rail";

const previousActEnvironment = Reflect.get(globalThis, "IS_REACT_ACT_ENVIRONMENT");
Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", true);
afterAll(() => {
  Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", previousActEnvironment);
  GlobalRegistrator.unregister();
});

// Keyboard coverage previously kept the heading list fixed. Exercise shrinking
// lists in both presentations, with the focused entry removed each time.
const items = Array.from(
  { length: 6 },
  (_, index) =>
    ({
      id: `heading-${index}`,
      label: `Heading ${index}`,
      level: 1,
    }) satisfies OutlineItem,
);

test("shrinking the outline keeps one valid tab stop and ArrowUp moves from it", async () => {
  const container = document.createElement("div");
  document.body.append(container);
  const root = createRoot(container);
  const scrollContainerRef = createRef<HTMLElement>();
  const render = async (presentation: "panel" | "rail", visibleItems: OutlineItem[]) => {
    const props = {
      items: visibleItems,
      scrollContainerRef,
      resolvePct: () => null,
      onJump: () => {},
      presentation,
    } satisfies FolioOutlineRailProps;
    await act(async () => root.render(<DefaultOutlineRail {...props} />));
  };

  try {
    for (const presentation of ["panel", "rail"] as const) {
      for (const remainingCount of [2, 3, 4]) {
        await render(presentation, items);
        const initialButtons = container.querySelectorAll("button");
        await act(async () => initialButtons.item(5)?.focus());

        const remainingItems = items.slice(0, remainingCount);
        await render(presentation, remainingItems);

        const list = container.querySelector("ol");
        const buttons = container.querySelectorAll("button");
        expect([...buttons].filter((button) => button.tabIndex === 0)).toHaveLength(1);
        expect(buttons.item(remainingCount - 1)?.tabIndex).toBe(0);

        await act(async () => {
          list?.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowUp", bubbles: true }));
        });

        expect(buttons.item(remainingCount - 2)?.tabIndex).toBe(0);
        expect(document.activeElement).toBe(buttons.item(remainingCount - 2));
        expect([...buttons].filter((button) => button.tabIndex === 0)).toHaveLength(1);
      }
    }
  } finally {
    await act(async () => root.unmount());
    container.remove();
  }
});

test("the outline panel renders only headings classified from the DOCX model", async () => {
  const headings = collectHeadings(
    HEADING_COLLECTOR_DOCUMENT,
    createBuiltInStyleIndex(HEADING_COLLECTOR_STYLES),
  );
  const classifiedItems = headings.map(
    ({ text, level }, index) =>
      ({ id: `heading-${index}`, label: text, level }) satisfies OutlineItem,
  );
  const container = document.createElement("div");
  document.body.append(container);
  const root = createRoot(container);
  const props = {
    items: classifiedItems,
    scrollContainerRef: createRef<HTMLElement>(),
    resolvePct: () => null,
    onJump: () => {},
    presentation: "panel",
  } satisfies FolioOutlineRailProps;

  try {
    await act(async () => root.render(<DefaultOutlineRail {...props} />));

    const labels = [...container.querySelectorAll(".folio-outline-item-label")].map(
      (label) => label.textContent ?? "",
    );
    expect(labels).toHaveLength(66);
    expect(labels).toEqual(headings.map(({ text }) => text));
    expect(labels).not.toContain(
      '"Synthetic Term" means a term used only by this synthetic document fixture.',
    );
    expect(labels).not.toContain("The parties agree that this ordinary clause remains body text.");
    expect(labels).not.toContain("Numbered list item at level 0");
    expect(labels).not.toContain("Numbered list item at level 1");
    expect(labels).not.toContain("Numbered list item at level 2");
    expect(labels).not.toContain("PLAIN BOLD CAPS");
  } finally {
    await act(async () => root.unmount());
    container.remove();
  }
});

// The former active-heading effect added another commit on every scroll update.
// Active headings now drive the outside tab stop directly; focus events own it inside.
test.each(["panel", "rail"] as const)(
  "%s follows active headings without another commit and preserves manual roving focus",
  async (presentation) => {
    const container = document.createElement("div");
    const outside = document.createElement("button");
    document.body.append(container, outside);
    const root = createRoot(container);
    const scrollContainerRef = createRef<HTMLElement>();
    let commits = 0;
    const profilerProps = { id: "outline", onRender: () => commits++ };
    const baseProps = {
      items,
      scrollContainerRef,
      resolvePct: () => null,
      onJump: () => {},
      presentation,
    } satisfies FolioOutlineRailProps;
    const render = async (activeId: string) => {
      await act(async () => {
        root.render(
          <Profiler {...profilerProps}>
            <DefaultOutlineRail {...baseProps} activeId={activeId} />
          </Profiler>,
        );
      });
    };
    const tabStop = () => container.querySelector<HTMLButtonElement>('button[tabindex="0"]');
    try {
      await render("heading-0");
      const buttons = container.querySelectorAll<HTMLButtonElement>("button");
      expect(tabStop()).toBe(buttons.item(0));
      const beforeScroll = commits;
      await render("heading-2");
      expect(commits - beforeScroll).toBe(1);
      expect(tabStop()).toBe(buttons.item(2));

      await act(async () => buttons.item(4)?.focus());
      const beforeFocusedScroll = commits;
      await render("heading-1");
      expect(commits - beforeFocusedScroll).toBe(1);
      expect(tabStop()).toBe(buttons.item(4));
      expect(document.activeElement).toBe(buttons.item(4));

      for (const [key, index] of [
        ["Home", 0],
        ["End", 5],
        ["ArrowUp", 4],
        ["ArrowDown", 5],
      ] as const) {
        await act(async () => {
          container
            .querySelector("ol")
            ?.dispatchEvent(new KeyboardEvent("keydown", { key, bubbles: true }));
        });
        expect(tabStop()).toBe(buttons.item(index));
        expect(document.activeElement).toBe(buttons.item(index));
      }
      await act(async () => outside.focus());
      expect(tabStop()).toBe(buttons.item(1));
      expect(document.activeElement).toBe(outside);
      expect(container.querySelectorAll('button[tabindex="0"]')).toHaveLength(1);
    } finally {
      await act(async () => root.unmount());
      container.remove();
      outside.remove();
    }
  },
);
