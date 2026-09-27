import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { afterAll, expect, test } from "bun:test";
import { act, createRef } from "react";
import { createRoot } from "react-dom/client";

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
