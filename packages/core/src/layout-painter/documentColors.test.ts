import { describe, expect, test } from "bun:test";

import {
  AUTHORED_BACKGROUND_COLOR_VAR,
  AUTHORED_TEXT_COLOR_VAR,
  getAutomaticTextColorForBackground,
  setAuthoredBackgroundColor,
  setAuthoredTextColor,
} from "./documentColors";

describe("document automatic text color", () => {
  test("uses black automatic text on explicit white document shading", () => {
    expect(getAutomaticTextColorForBackground("#FFFFFF")).toBe("#000000");
  });

  test("uses white automatic text on explicit dark document shading", () => {
    expect(getAutomaticTextColorForBackground("#111111")).toBe("#FFFFFF");
  });

  test("uses black automatic text on mid-tone document shading", () => {
    expect(getAutomaticTextColorForBackground("#A9A9A9")).toBe("#000000");
  });

  test("leaves automatic text theme-adaptive when shading is not a concrete color", () => {
    expect(getAutomaticTextColorForBackground("auto")).toBeUndefined();
    expect(getAutomaticTextColorForBackground(undefined)).toBeUndefined();
  });
});

describe("authored document backgrounds", () => {
  test("retains the source color for stylesheet dark-mode adaptation", () => {
    const values: Record<string, string> = {};
    const style = {
      backgroundColor: "",
      setProperty: (name: string, value: string) => {
        values[name] = value;
      },
    } as CSSStyleDeclaration;

    setAuthoredBackgroundColor(style, "#F8F2EB");

    expect(style.backgroundColor).toBe("#F8F2EB");
    expect(values[AUTHORED_BACKGROUND_COLOR_VAR]).toBe("#F8F2EB");
  });
});

describe("authored document text", () => {
  test("retains the source color for stylesheet dark-mode adaptation", () => {
    const values: Record<string, string> = {};
    const style = {
      color: "",
      setProperty: (name: string, value: string) => {
        values[name] = value;
      },
    } as CSSStyleDeclaration;

    setAuthoredTextColor(style, "#000000");

    expect(style.color).toBe("#000000");
    expect(values[AUTHORED_TEXT_COLOR_VAR]).toBe("#000000");
  });
});
