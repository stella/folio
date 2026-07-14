import { describe, expect, test } from "bun:test";

import type { Theme } from "../../types/document";
import { resolveEffectiveTableCellFormatting } from "./effectiveTableCellFormatting";

type ResolveOptions = Parameters<typeof resolveEffectiveTableCellFormatting>[0];

const resolveFormatting = (overrides: Partial<ResolveOptions>) =>
  resolveEffectiveTableCellFormatting({
    directFormatting: undefined,
    styleFormatting: undefined,
    tableBorders: undefined,
    position: {},
    gridWidthPercent: undefined,
    defaultMargins: undefined,
    theme: undefined,
    ...overrides,
  });

const theme: Theme = {
  colorScheme: {
    dk1: "000000",
    lt1: "FFFFFF",
    dk2: "44546A",
    lt2: "E7E6E6",
    accent1: "4472C4",
    accent2: "ED7D31",
    accent3: "A5A5A5",
    accent4: "FFC000",
    accent5: "5B9BD5",
    accent6: "70AD47",
    hlink: "0563C1",
    folHlink: "954F72",
  },
};

describe("resolveEffectiveTableCellFormatting", () => {
  test("keeps a direct no-fill distinct from an absent or inherited background", () => {
    const result = resolveFormatting({
      directFormatting: { shading: { pattern: "nil" } },
      styleFormatting: {
        shading: { pattern: "clear", fill: { themeColor: "accent1" } },
      },
      theme,
    });

    expect(result.background).toEqual({ type: "none", source: "direct" });
  });

  test("records whether width came from the authored cell or table grid", () => {
    const direct = resolveFormatting({
      directFormatting: { width: { value: 1440, type: "dxa" } },
      gridWidthPercent: 50,
    });
    const grid = resolveFormatting({ gridWidthPercent: 50 });

    expect(direct.width).toEqual({
      type: "value",
      source: "direct",
      value: 1440,
      widthType: "dxa",
    });
    expect(grid.width).toEqual({
      type: "value",
      source: "grid",
      value: 50,
      widthType: "pct",
    });
  });

  test("applies direct, style, and table border precedence before resolving theme colors", () => {
    const result = resolveFormatting({
      directFormatting: {
        borders: { left: { style: "dashed", color: { rgb: "123456" } } },
      },
      styleFormatting: {
        borders: { top: { style: "double", color: { themeColor: "accent2" } } },
      },
      tableBorders: {
        top: { style: "single", color: { themeColor: "accent1" } },
        bottom: { style: "single", color: { themeColor: "accent1" } },
      },
      position: { isFirstRow: true, isLastRow: true },
      theme,
    });

    expect(result.borders?.top).toEqual({ style: "double", color: { rgb: "ED7D31" } });
    expect(result.borders?.bottom).toEqual({ style: "single", color: { rgb: "4472C4" } });
    expect(result.borders?.left).toEqual({ style: "dashed", color: { rgb: "123456" } });
  });

  test("uses direct margins before style and table defaults", () => {
    const result = resolveFormatting({
      directFormatting: { margins: { left: { value: 120, type: "dxa" } } },
      styleFormatting: { margins: { left: { value: 240, type: "dxa" } } },
      defaultMargins: { left: 360 },
    });

    expect(result.margins).toEqual({ left: 120 });
  });
});
