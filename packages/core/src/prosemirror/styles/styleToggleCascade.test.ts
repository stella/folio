import { describe, expect, test } from "bun:test";

import type { TextFormatting } from "../../types/document";
import { cascadeStyleTextFormatting } from "./styleToggleCascade";

const toggleKeys = [
  "bold",
  "boldCs",
  "italic",
  "italicCs",
  "allCaps",
  "emboss",
  "imprint",
  "outline",
  "shadow",
  "smallCaps",
  "strike",
  "hidden",
] as const satisfies readonly (keyof TextFormatting)[];

describe("style toggle cascade", () => {
  for (const key of toggleKeys) {
    test(`${key} reverses once per style level`, () => {
      const lower: TextFormatting = { [key]: true };
      const upper: TextFormatting = { [key]: true };

      const result = cascadeStyleTextFormatting([
        { formatting: lower, type: "style" },
        { formatting: upper, type: "style" },
      ]);

      expect(result.formatting?.[key]).toBe(false);
    });

    test(`${key} distinguishes defaults, style, and direct false`, () => {
      const enabled: TextFormatting = { [key]: true };
      const disabled: TextFormatting = { [key]: false };

      expect(
        cascadeStyleTextFormatting([
          { formatting: enabled, type: "defaults" },
          { formatting: disabled, type: "style" },
        ]).formatting?.[key],
      ).toBe(true);
      expect(
        cascadeStyleTextFormatting([
          { formatting: enabled, type: "style" },
          { formatting: disabled, type: "style" },
        ]).formatting?.[key],
      ).toBe(true);
      expect(
        cascadeStyleTextFormatting([
          { formatting: enabled, type: "style" },
          { formatting: disabled, type: "direct" },
        ]).formatting?.[key],
      ).toBe(false);
    });
  }

  test("style off behaves like omission", () => {
    const inherited = cascadeStyleTextFormatting([{ formatting: { bold: true }, type: "style" }]);
    const absent = cascadeStyleTextFormatting([
      { cascade: inherited, type: "carried" },
      { formatting: { italic: true }, type: "style" },
    ]);
    const off = cascadeStyleTextFormatting([
      { cascade: inherited, type: "carried" },
      { formatting: { bold: false }, type: "style" },
    ]);

    expect(absent.formatting?.bold).toBe(true);
    expect(off.formatting?.bold).toBe(true);
  });

  test("direct formatting sets rather than toggles", () => {
    const inherited = cascadeStyleTextFormatting([
      { formatting: { bold: true }, type: "style" },
      { formatting: { bold: true }, type: "style" },
    ]);

    expect(
      cascadeStyleTextFormatting([
        { cascade: inherited, type: "carried" },
        { formatting: { bold: true }, type: "direct" },
      ]).formatting?.bold,
    ).toBe(true);
    expect(
      cascadeStyleTextFormatting([
        { cascade: inherited, type: "carried" },
        { formatting: { bold: false }, type: "direct" },
      ]).formatting?.bold,
    ).toBe(false);
  });

  test("enabled document defaults retain their priority over style toggles", () => {
    const result = cascadeStyleTextFormatting([
      { formatting: { bold: true }, type: "defaults" },
      { formatting: { bold: true }, type: "style" },
      { formatting: { bold: true }, type: "style" },
    ]);

    expect(result.formatting?.bold).toBe(true);
  });

  test("cloning between paragraph and character passes preserves toggle history", () => {
    const paragraph = cascadeStyleTextFormatting([
      { formatting: { bold: true }, type: "defaults" },
      { formatting: { bold: true }, type: "style" },
    ]);
    const clonedParagraph = {
      ...paragraph,
      formatting: paragraph.formatting ? { ...paragraph.formatting } : undefined,
      toggleStates: new Map(paragraph.toggleStates),
    };
    expect(clonedParagraph.formatting).not.toBe(paragraph.formatting);
    expect(clonedParagraph.toggleStates).not.toBe(paragraph.toggleStates);
    const character = cascadeStyleTextFormatting([
      { cascade: clonedParagraph, type: "carried" },
      { formatting: { bold: true }, type: "style" },
    ]);

    expect(character.formatting?.bold).toBe(true);
  });

  test("an explicit style off preserves the document-default priority", () => {
    const result = cascadeStyleTextFormatting([
      { formatting: { bold: true }, type: "defaults" },
      { formatting: { bold: false }, type: "style" },
      { formatting: { bold: true }, type: "style" },
    ]);

    expect(result.formatting?.bold).toBe(true);
  });

  test("direct formatting clears the document-default priority", () => {
    const inherited = cascadeStyleTextFormatting([
      { formatting: { bold: true }, type: "defaults" },
      { formatting: { bold: true }, type: "style" },
    ]);
    const result = cascadeStyleTextFormatting([
      { cascade: inherited, type: "carried" },
      { formatting: { bold: false }, type: "direct" },
    ]);

    expect(result.formatting?.bold).toBe(false);
  });

  test("double strike keeps ordinary last-defined inheritance", () => {
    const result = cascadeStyleTextFormatting([
      { formatting: { doubleStrike: true }, type: "style" },
      { formatting: { doubleStrike: true }, type: "style" },
    ]);

    expect(result.formatting?.doubleStrike).toBe(true);
    expect(
      cascadeStyleTextFormatting([
        { formatting: { doubleStrike: true }, type: "style" },
        { formatting: { doubleStrike: false }, type: "style" },
      ]).formatting?.doubleStrike,
    ).toBe(false);
  });
});
