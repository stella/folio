import { describe, expect, test } from "bun:test";

import {
  hasRunFormattingOverrideAttrs,
  withAuthoredRunFormatting,
} from "./runFormattingProvenance";

describe("run formatting override ownership", () => {
  test("retains sparse known-empty authored baselines", () => {
    expect(hasRunFormattingOverrideAttrs({ _authoredOn: [] })).toBe(true);
    expect(hasRunFormattingOverrideAttrs({ _authoredOff: [] })).toBe(true);
    expect(
      hasRunFormattingOverrideAttrs(withAuthoredRunFormatting({ bold: false }, undefined)),
    ).toBe(true);
  });

  test("drops empty rendering-only metadata", () => {
    expect(hasRunFormattingOverrideAttrs({})).toBe(false);
    expect(hasRunFormattingOverrideAttrs({ _authoredValues: {} })).toBe(false);
    expect(hasRunFormattingOverrideAttrs({ directFontProperties: [] })).toBe(false);
    expect(hasRunFormattingOverrideAttrs({ complexScriptPropertyAbsences: [] })).toBe(false);
  });

  test("retains every non-empty authorship and rendering signal class", () => {
    expect(hasRunFormattingOverrideAttrs({ _authoredOn: ["bold"] })).toBe(true);
    expect(hasRunFormattingOverrideAttrs({ _authoredOff: ["italic"] })).toBe(true);
    expect(hasRunFormattingOverrideAttrs({ _authoredValues: { fontSize: 24 } })).toBe(true);
    expect(hasRunFormattingOverrideAttrs({ directFontProperties: ["color"] })).toBe(true);
    expect(hasRunFormattingOverrideAttrs({ complexScriptPropertyAbsences: ["boldCs"] })).toBe(true);
    expect(hasRunFormattingOverrideAttrs({ bold: false })).toBe(true);
    expect(hasRunFormattingOverrideAttrs({ color: "auto" })).toBe(true);
    expect(hasRunFormattingOverrideAttrs({ shading: { pattern: "nil" } })).toBe(true);
  });
});
