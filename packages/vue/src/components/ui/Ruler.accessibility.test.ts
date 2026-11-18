import { compileScript, parse } from "@vue/compiler-sfc";
import { panic } from "better-result";
import { plugin } from "bun";
import { afterAll, expect, test } from "bun:test";

import { acquireHarnessDom } from "../../../../../test/harnessDom";

const releaseDom = acquireHarnessDom();
afterAll(releaseDom);

// Exercise the actual SFC templates and event bindings, without a browser build.
plugin({
  name: "ruler-sfc-test",
  setup(build) {
    build.onLoad(
      { filter: /\/(Horizontal|Vertical)Ruler\.vue$/ },
      async ({ path }) => {
        const { descriptor } = parse(await Bun.file(path).text(), {
          filename: path,
        });
        const compiled = compileScript(descriptor, {
          id: path,
          inlineTemplate: true,
        });
        return { contents: compiled.content, loader: "ts" };
      },
    );
  },
});

const { createApp, defineComponent, h, nextTick, reactive } =
  await import("vue");
const { default: HorizontalRuler } = await import("./HorizontalRuler.vue");
const { default: VerticalRuler } = await import("./VerticalRuler.vue");

test.each([0.5, 2])(
  "all Vue ruler handles expose values and edit through keys at zoom %s",
  async (zoom) => {
    const container = document.createElement("div");
    document.body.append(container);
    const state = reactive({ editable: true, unit: "inch" as "inch" | "cm" });
    const changes: [string, number][] = [];
    const change = (name: string) => (value: number) =>
      changes.push([name, value]);
    const app = createApp(
      defineComponent({
        setup() {
          return () =>
            h("div", [
              h(HorizontalRuler, {
                zoom,
                unit: state.unit,
                editable: state.editable,
                showFirstLineIndent: true,
                hangingIndent: true,
                firstLineIndent: 360,
                indentLeft: 720,
                indentRight: 360,
                onLeftMarginChange: change("leftMargin"),
                onRightMarginChange: change("rightMargin"),
                onFirstLineIndentChange: change("firstLine"),
                onIndentLeftChange: change("leftIndent"),
                onIndentRightChange: change("rightIndent"),
              }),
              h(VerticalRuler, {
                zoom,
                unit: state.unit,
                editable: state.editable,
                onTopMarginChange: change("topMargin"),
                onBottomMarginChange: change("bottomMargin"),
              }),
            ]);
        },
      }),
    );
    app.mount(container);
    try {
      for (const unit of ["inch", "cm"] as const) {
        state.unit = unit;
        await nextTick();
        const sliders =
          container.querySelectorAll<HTMLElement>('[role="slider"]');
        expect(sliders.length).toBe(7);
        expect(container.querySelectorAll('[role="group"]').length).toBe(2);
        const expected = [
          ["leftMargin", 1440, 0, 10080],
          ["rightMargin", 1440, 0, 10080],
          ["firstLine", -360, -720, 7560],
          ["leftIndent", 720, 0, 8280],
          ["rightIndent", 360, 0, 7920],
          ["topMargin", 1440, 0, 13680],
          ["bottomMargin", 1440, 0, 13680],
        ] as const;
        for (const [index, entry] of expected.entries()) {
          const slider = sliders.item(index);
          const [name, value, min, max] = entry;
          expect(slider.getAttribute("aria-label")).toBeTruthy();
          expect(slider.tabIndex).toBe(0);
          expect(slider.getAttribute("aria-valuenow")).toBe(String(value));
          expect(slider.getAttribute("aria-valuemin")).toBe(String(min));
          expect(slider.getAttribute("aria-valuemax")).toBe(String(max));
          expect(slider.getAttribute("aria-valuetext")).toContain(
            unit === "inch" ? "in" : "cm",
          );
          slider.focus();
          expect(document.activeElement).toBe(slider);
          for (const [key, next] of [
            ["ArrowUp", Math.round(value + (unit === "inch" ? 180 : 56.7))],
            ["Home", min],
            ["End", max],
          ] as const) {
            const event = new KeyboardEvent("keydown", {
              key,
              bubbles: true,
              cancelable: true,
            });
            slider.dispatchEvent(event);
            expect(event.defaultPrevented).toBe(true);
            expect(changes.at(-1)).toEqual([name, next]);
          }

          // Extreme pointer positions and Home/End reach the same limits. This
          // exercises adapter coordinate conversion as well as shared bounds.
          const pointerValues: number[] = [];
          slider.dispatchEvent(
            new MouseEvent("mousedown", {
              bubbles: true,
              clientX: 0,
              clientY: 0,
            }),
          );
          for (const position of [-100000, 100000]) {
            document.dispatchEvent(
              new MouseEvent("mousemove", {
                clientX: position,
                clientY: position,
              }),
            );
            pointerValues.push(changes.at(-1)?.[1] ?? NaN);
          }
          document.dispatchEvent(new MouseEvent("mouseup"));
          await nextTick();
          expect(pointerValues.toSorted((a, b) => a - b)).toEqual([min, max]);
        }
      }
      state.editable = false;
      await nextTick();
      expect(container.querySelectorAll('[role="slider"]').length).toBe(0);
      expect(container.querySelector('[tabindex="0"]')).toBeNull();
    } finally {
      app.unmount();
      container.remove();
    }
  },
);

test("Vue keyboard edits follow controlled authored values and updated page bounds", async () => {
  const container = document.createElement("div");
  document.body.append(container);
  const state = reactive({ indentLeft: -720, marginRight: 1440 });
  const app = createApp(
    defineComponent({
      setup() {
        return () =>
          h(HorizontalRuler, {
            editable: true,
            indentLeft: state.indentLeft,
            sectionProps: { marginRight: state.marginRight },
            onIndentLeftChange: (value: number) => {
              state.indentLeft = value;
            },
          });
      },
    }),
  );
  app.mount(container);
  try {
    const slider = container.querySelector<HTMLElement>(
      '[aria-label="Left indent"]',
    );
    expect(slider).not.toBeNull();
    if (!slider) panic("Expected a mounted ruler slider");
    slider.focus();
    slider.dispatchEvent(
      new KeyboardEvent("keydown", { key: "ArrowRight", bubbles: true }),
    );
    expect(state.indentLeft).toBe(-540);
    state.marginRight = 5000;
    await nextTick();
    expect(slider.getAttribute("aria-valuenow")).toBe("-540");
    expect(slider.getAttribute("aria-valuemin")).toBe("-540");
    expect(slider.getAttribute("aria-valuemax")).toBe("5080");
    expect(document.activeElement).toBe(slider);
    slider.dispatchEvent(
      new KeyboardEvent("keydown", { key: "ArrowRight", bubbles: true }),
    );
    expect(state.indentLeft).toBe(-360);
    await nextTick();
    expect(document.activeElement).toBe(slider);
    slider.dispatchEvent(
      new KeyboardEvent("keydown", { key: "End", bubbles: true }),
    );
    expect(state.indentLeft).toBe(5080);
  } finally {
    app.unmount();
    container.remove();
  }
});
