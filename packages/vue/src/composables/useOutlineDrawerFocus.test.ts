import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { afterAll, expect, test } from "bun:test";
import type { Ref } from "vue";

const { createApp, defineComponent, h, nextTick, ref } = await import("vue");
const { useOutlineDrawerFocus } = await import("./useOutlineDrawerFocus");

afterAll(() => GlobalRegistrator.unregister());

test("an available drawer focuses and closes with Escape when no headings survive filtering", async () => {
  const container = document.createElement("div");
  document.body.append(container);
  const state: {
    available?: Ref<boolean>;
    surface?: Ref<"column" | "rail" | "drawer">;
    closeCount: number;
  } = { closeCount: 0 };
  const app = createApp(
    defineComponent({
      setup() {
        const available = ref(false);
        const surface = ref<"column" | "rail" | "drawer">("drawer");
        const navRef = ref<HTMLElement | null>(null);
        const headings: string[] = [];
        state.available = available;
        state.surface = surface;
        useOutlineDrawerFocus({
          available,
          surface,
          navRef,
          onClose: () => {
            state.closeCount++;
            surface.value = "rail";
          },
        });
        return () =>
          h("div", [
            h("button", { id: "outline-trigger" }, "Outline"),
            surface.value === "drawer" && available.value
              ? h("nav", { ref: navRef }, [
                  headings.length >= 2 ? h("ol", { id: "headings" }) : null,
                  h("button", { id: "drawer-close" }, "Close"),
                ])
              : null,
          ]);
      },
    }),
  );
  app.mount(container);

  try {
    const opener = container.querySelector<HTMLButtonElement>("#outline-trigger");
    if (!opener) throw new Error("outline trigger missing");
    opener.focus();
    expect(container.querySelector("nav")).toBeNull();
    const available = state.available;
    const surface = state.surface;
    if (!available || !surface) throw new Error("drawer state missing");

    available.value = true;
    await nextTick();
    await nextTick();

    const drawerAction = container.querySelector<HTMLButtonElement>("#drawer-close");
    expect(container.querySelector("#headings")).toBeNull();
    expect(document.activeElement).toBe(drawerAction);

    const escape = new KeyboardEvent("keydown", {
      key: "Escape",
      bubbles: true,
      cancelable: true,
    });
    document.dispatchEvent(escape);
    await nextTick();

    expect(escape.defaultPrevented).toBe(true);
    expect(state.closeCount).toBe(1);
    expect(surface.value).toBe("rail");
    expect(document.activeElement).toBe(opener);
  } finally {
    app.unmount();
    container.remove();
  }
});
