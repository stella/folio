import { describe, expect, test } from "bun:test";

import { isValidHexColor, resolveColor, resolveColorToHex } from "./colorResolver";

// A crafted `w:val`/`w:fill` (or a themeN.xml color-scheme swatch) that
// smuggles a CSS `url()` payload instead of a real color — this is the
// shape a themed table-fill / diagonal-border CSS-injection finding takes.
const MALICIOUS_COLOR = "FFFFFF 0,red 1px),url(//x)";

describe("isValidHexColor", () => {
  test("accepts 3/6/8-digit hex, with or without a leading #", () => {
    expect(isValidHexColor("FF0000")).toBe(true);
    expect(isValidHexColor("#FF0000")).toBe(true);
    expect(isValidHexColor("F00")).toBe(true);
    expect(isValidHexColor("FF0000AA")).toBe(true);
  });

  test("rejects a value carrying a CSS url() payload", () => {
    expect(isValidHexColor(MALICIOUS_COLOR)).toBe(false);
  });

  test("rejects empty/undefined/null", () => {
    expect(isValidHexColor("")).toBe(false);
    expect(isValidHexColor(undefined)).toBe(false);
    expect(isValidHexColor(null)).toBe(false);
  });
});

describe("resolveColor", () => {
  test("resolves a valid rgb color", () => {
    expect(resolveColor({ rgb: "ff0000" }, null)).toBe("#FF0000");
  });

  test("falls back to the safe default when rgb carries a CSS injection payload", () => {
    expect(resolveColor({ rgb: MALICIOUS_COLOR }, null)).toBe("#000000");
    expect(resolveColor({ rgb: MALICIOUS_COLOR }, null, "FF0000")).toBe("#FF0000");
  });

  test("falls back to the safe default when a theme-resolved swatch is malformed", () => {
    const maliciousTheme = {
      colorScheme: {
        dk1: "000000",
        lt1: "FFFFFF",
        dk2: "44546A",
        lt2: "E7E6E6",
        accent1: MALICIOUS_COLOR,
        accent2: "ED7D31",
        accent3: "A5A5A5",
        accent4: "FFC000",
        accent5: "5B9BD5",
        accent6: "70AD47",
        hlink: "0563C1",
        folHlink: "954F72",
      },
    };
    expect(resolveColor({ themeColor: "accent1" }, maliciousTheme)).toBe("#000000");
  });

  test("auto and undefined color both resolve to the default", () => {
    expect(resolveColor({ auto: true }, null)).toBe("#000000");
    expect(resolveColor(undefined, null)).toBe("#000000");
  });
});

describe("resolveColorToHex", () => {
  test("returns the normalized hex for a valid rgb color", () => {
    expect(resolveColorToHex({ rgb: "ff0000" }, null)).toBe("FF0000");
  });

  test("returns undefined (safe fallback / no-color) for a CSS injection payload", () => {
    expect(resolveColorToHex({ rgb: MALICIOUS_COLOR }, null)).toBeUndefined();
  });

  test("returns undefined for auto/absent colors", () => {
    expect(resolveColorToHex({ auto: true }, null)).toBeUndefined();
    expect(resolveColorToHex(undefined, null)).toBeUndefined();
  });
});
