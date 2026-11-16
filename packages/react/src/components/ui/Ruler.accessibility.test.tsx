import { act } from "react";
import { createRoot } from "react-dom/client";

import { panic } from "better-result";
import { afterAll, expect, test } from "bun:test";
import { IntlProvider } from "use-intl";

import { getFolioMessages } from "@stll/folio-core/i18n/messages";

import { acquireHarnessDom } from "../../../../../test/harnessDom";
import { HorizontalRuler } from "./HorizontalRuler";
import { VerticalRuler } from "./VerticalRuler";

const INITIAL_SECTION_PROPS = { marginRight: 1440 };
const NARROW_SECTION_PROPS = { marginRight: 5000 };

const releaseDom = acquireHarnessDom();
const previousActEnvironment = Reflect.get(
  globalThis,
  "IS_REACT_ACT_ENVIRONMENT",
);
Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", true);
afterAll(() => {
  Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", previousActEnvironment);
  releaseDom();
});

test.each([0.5, 2])(
  "all React ruler handles expose values and edit through keys at zoom %s",
  async (zoom) => {
    const container = document.createElement("div");
    document.body.append(container);
    const root = createRoot(container);
    const changes: [string, number][] = [];
    const change = (name: string) => (value: number) =>
      changes.push([name, value]);
    async function render(editable: boolean, unit: "inch" | "cm") {
      await act(async () =>
        root.render(
          <IntlProvider
            locale="en"
            timeZone="UTC"
            messages={getFolioMessages("en")}
          >
            <HorizontalRuler
              zoom={zoom}
              unit={unit}
              editable={editable}
              showFirstLineIndent
              hangingIndent
              firstLineIndent={360}
              indentLeft={720}
              indentRight={360}
              onLeftMarginChange={change("leftMargin")}
              onRightMarginChange={change("rightMargin")}
              onFirstLineIndentChange={change("firstLine")}
              onIndentLeftChange={change("leftIndent")}
              onIndentRightChange={change("rightIndent")}
            />
            <VerticalRuler
              zoom={zoom}
              unit={unit}
              editable={editable}
              onTopMarginChange={change("topMargin")}
              onBottomMarginChange={change("bottomMargin")}
            />
          </IntlProvider>,
        ),
      );
    }
    try {
      for (const unit of ["inch", "cm"] as const) {
        await render(true, unit);
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
          await act(async () => {
            slider.dispatchEvent(
              new MouseEvent("mousedown", {
                bubbles: true,
                clientX: 0,
                clientY: 0,
              }),
            );
          });
          for (const position of [-100000, 100000]) {
            await act(async () => {
              document.dispatchEvent(
                new MouseEvent("mousemove", {
                  clientX: position,
                  clientY: position,
                }),
              );
            });
            pointerValues.push(changes.at(-1)?.[1] ?? NaN);
          }
          await act(async () => {
            document.dispatchEvent(new MouseEvent("mouseup"));
          });
          expect(pointerValues.toSorted((a, b) => a - b)).toEqual([min, max]);
        }
      }
      await render(false, "inch");
      expect(container.querySelectorAll('[role="slider"]').length).toBe(0);
      expect(container.querySelector('[tabindex="0"]')).toBeNull();
    } finally {
      await act(async () => root.unmount());
      container.remove();
    }
  },
);

test("React keyboard edits follow controlled values and updated page bounds; missing callbacks create no controls", async () => {
  const container = document.createElement("div");
  document.body.append(container);
  const root = createRoot(container);
  let indentLeft = -720;
  let sectionProps = INITIAL_SECTION_PROPS;
  const change = (value: number) => {
    indentLeft = value;
  };
  async function render(withCallback: boolean) {
    await act(async () =>
      root.render(
        <IntlProvider
          locale="en"
          timeZone="UTC"
          messages={getFolioMessages("en")}
        >
          <HorizontalRuler
            editable
            showFirstLineIndent
            indentLeft={indentLeft}
            sectionProps={sectionProps}
            onIndentLeftChange={withCallback ? change : undefined}
          />
        </IntlProvider>,
      ),
    );
  }
  try {
    await render(true);
    const slider = container.querySelector<HTMLElement>('[role="slider"]');
    expect(slider).not.toBeNull();
    if (!slider) panic("Expected a mounted ruler slider");
    slider.focus();
    slider.dispatchEvent(
      new KeyboardEvent("keydown", { key: "ArrowRight", bubbles: true }),
    );
    expect(indentLeft).toBe(-540);
    sectionProps = NARROW_SECTION_PROPS;
    await render(true);
    expect(slider.getAttribute("aria-valuenow")).toBe("-540");
    expect(slider.getAttribute("aria-valuemin")).toBe("-540");
    expect(slider.getAttribute("aria-valuemax")).toBe("5080");
    expect(document.activeElement).toBe(slider);
    slider.dispatchEvent(
      new KeyboardEvent("keydown", { key: "ArrowRight", bubbles: true }),
    );
    expect(indentLeft).toBe(-360);
    await render(true);
    expect(document.activeElement).toBe(slider);
    slider.dispatchEvent(
      new KeyboardEvent("keydown", { key: "End", bubbles: true }),
    );
    expect(indentLeft).toBe(5080);
    await render(false);
    expect(container.querySelector('[role="slider"]')).toBeNull();
  } finally {
    await act(async () => root.unmount());
    container.remove();
  }
});
