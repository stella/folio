import { afterAll, expect, test } from "bun:test";

import { acquireHarnessDom } from "../../../../test/harnessDom";

const releaseDom = acquireHarnessDom();
afterAll(releaseDom);

const { createApp, h, nextTick, ref } = await import("vue");
const { useFixedDropdown } = await import("./useFixedDropdown");

test("a canceled editor pointer press closes an open dropdown", async () => {
  const container = document.createElement("div");
  const editor = document.createElement("div");
  document.body.append(container, editor);
  const isOpen = ref(false);
  const app = createApp({
    setup() {
      const dropdown = useFixedDropdown({
        isOpen,
        onClose: () => {
          isOpen.value = false;
        },
      });
      return () =>
        h("div", [
          h("button", { ref: dropdown.containerRef }, "trigger"),
          h("div", { ref: dropdown.dropdownRef }, "menu"),
        ]);
    },
  });
  editor.addEventListener("pointerdown", (event) => {
    event.preventDefault();
    event.stopPropagation();
  });
  app.mount(container);
  try {
    isOpen.value = true;
    await nextTick();
    editor.dispatchEvent(
      new PointerEvent("pointerdown", {
        bubbles: true,
        cancelable: true,
      }),
    );
    expect(isOpen.value).toBe(false);
  } finally {
    app.unmount();
    container.remove();
    editor.remove();
  }
});
