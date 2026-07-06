import { describe, expect, test } from "bun:test";
import { defineComponent } from "vue";

import { DEFAULT_COMPONENTS, resolveFolioComponents } from "./folio-ui";

// A sentinel override component; identity is all the resolution test checks.
const InjectedColorPicker = defineComponent({ name: "InjectedColorPicker", render: () => null });

describe("resolveFolioComponents", () => {
  test("returns the defaults when nothing is injected", () => {
    expect(resolveFolioComponents()).toBe(DEFAULT_COMPONENTS);
    expect(resolveFolioComponents(undefined)).toBe(DEFAULT_COMPONENTS);
  });

  test("an empty override still resolves every contract key to a default", () => {
    // A non-undefined override forces a fresh merged object, so identity differs
    // but every value equals its default.
    expect(resolveFolioComponents({})).toEqual(DEFAULT_COMPONENTS);
  });

  test("an injected component wins; unspecified keys fall back to defaults", () => {
    const resolved = resolveFolioComponents({ ColorPicker: InjectedColorPicker });
    expect(resolved.ColorPicker).toBe(InjectedColorPicker);
    // Every other key is untouched.
    expect(resolved.Button).toBe(DEFAULT_COMPONENTS.Button);
    expect(resolved.Popover).toBe(DEFAULT_COMPONENTS.Popover);
    expect(resolved.Menu).toBe(DEFAULT_COMPONENTS.Menu);
    expect(resolved.Dialog).toBe(DEFAULT_COMPONENTS.Dialog);
    expect(resolved.Select).toBe(DEFAULT_COMPONENTS.Select);
    expect(resolved.Input).toBe(DEFAULT_COMPONENTS.Input);
    expect(resolved.Checkbox).toBe(DEFAULT_COMPONENTS.Checkbox);
    expect(resolved.DatePickerPopover).toBe(DEFAULT_COMPONENTS.DatePickerPopover);
    expect(resolved.OutlineRail).toBe(DEFAULT_COMPONENTS.OutlineRail);
  });

  test("the full ten-primitive contract has a distinct default per key", () => {
    const keys = Object.keys(DEFAULT_COMPONENTS);
    expect(keys.length).toBe(10);
    expect(new Set(Object.values(DEFAULT_COMPONENTS)).size).toBe(10);
  });

  test("resolution does not mutate the shared defaults object", () => {
    const before = { ...DEFAULT_COMPONENTS };
    resolveFolioComponents({ ColorPicker: InjectedColorPicker });
    expect(DEFAULT_COMPONENTS).toEqual(before);
    expect(DEFAULT_COMPONENTS.ColorPicker).not.toBe(InjectedColorPicker);
  });

  test("every contract key has a non-null default (no gaps standalone)", () => {
    for (const component of Object.values(DEFAULT_COMPONENTS)) {
      expect(component).toBeDefined();
    }
  });
});
