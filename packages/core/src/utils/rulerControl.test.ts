import { expect, mock, test } from "bun:test";

import { createRulerControl } from "./rulerControl";

function keyEvent(key: string) {
  return {
    key,
    altKey: false,
    ctrlKey: false,
    metaKey: false,
    preventDefault: mock(() => {}),
    stopPropagation: mock(() => {}),
  };
}

test.each(["inch", "cm"] as const)(
  "%s keys share announced bounds, including authored values outside pointer bounds",
  (unit) => {
    for (const value of [-720, 0, 300, 720, 1800]) {
      const onChange = mock((_value: number) => {});
      const control = createRulerControl({
        value,
        min: 0,
        max: 1440,
        unit,
        onChange,
      });
      const min = Math.min(0, value);
      const max = Math.max(1440, value);
      expect(control.attributes["aria-valuenow"]).toBe(value);
      expect(control.attributes["aria-valuemin"]).toBe(min);
      expect(control.attributes["aria-valuemax"]).toBe(max);
      for (const key of [
        "ArrowUp",
        "ArrowRight",
        "ArrowDown",
        "ArrowLeft",
        "PageUp",
        "PageDown",
        "Home",
        "End",
      ]) {
        onChange.mockClear();
        const event = keyEvent(key);
        control.onKeyDown(event);
        expect(event.preventDefault).toHaveBeenCalledTimes(1);
        expect(event.stopPropagation).toHaveBeenCalledTimes(1);
        const next = onChange.mock.calls.at(0)?.at(0) ?? value;
        expect(Number.isInteger(next)).toBe(true);
        expect(next).toBeGreaterThanOrEqual(min);
        expect(next).toBeLessThanOrEqual(max);
        if (key === "Home") expect(next).toBe(min);
        if (key === "End") expect(next).toBe(max);
      }
      onChange.mockClear();
      control.onKeyDown(keyEvent("ArrowUp"));
      expect(onChange.mock.calls.at(0)?.at(0) ?? value).toBe(
        Math.min(max, Math.round(value + (unit === "inch" ? 180 : 56.7))),
      );
    }
  },
);

test("inactive controls and modified/unrelated keys leave events and values untouched", () => {
  const onChange = mock((_value: number) => {});
  const active = createRulerControl({
    value: 720,
    min: -720,
    max: 1440,
    unit: "inch",
    onChange,
  });
  const inactive = createRulerControl({
    value: 720,
    min: 0,
    max: 1440,
    unit: "inch",
  });
  expect(inactive.attributes.role).toBeUndefined();
  expect(inactive.attributes.tabIndex).toBeUndefined();
  for (const event of [
    keyEvent("Tab"),
    { ...keyEvent("ArrowRight"), ctrlKey: true },
    { ...keyEvent("End"), metaKey: true },
    { ...keyEvent("Home"), altKey: true },
  ]) {
    active.onKeyDown(event);
    inactive.onKeyDown(event);
    expect(event.preventDefault).not.toHaveBeenCalled();
  }
  const event = keyEvent("ArrowUp");
  inactive.onKeyDown(event);
  expect(event.preventDefault).not.toHaveBeenCalled();
  expect(onChange).not.toHaveBeenCalled();
});
