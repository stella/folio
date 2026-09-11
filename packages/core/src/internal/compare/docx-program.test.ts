import { describe, expect, test } from "bun:test";

import {
  DocxComparisonProgram,
  type DocxComparisonInstructionInput,
} from "./docx-program";

const paragraphProperties = Object.freeze({
  styleId: null,
  listLevel: null,
  directAlignment: null,
  directSpacing: null,
});

const source = Object.freeze({
  blockId: "p-1",
  kind: "paragraph",
  text: "old",
  paragraphProperties,
  structuralBoundaries: [],
  containerPath: [],
});

const replacement = (): Extract<
  DocxComparisonInstructionInput,
  { readonly type: "replaceText" }
> => ({
  type: "replaceText",
  source,
  sourceStartOffset: 0,
  range: {
    sourceText: "old",
    targetText: "new",
    segments: [
      {
        type: "del",
        text: "old",
        baseStart: 0,
        baseEnd: 3,
        revisedStart: 0,
        revisedEnd: 0,
      },
      {
        type: "ins",
        text: "new",
        baseStart: 3,
        baseEnd: 3,
        revisedStart: 0,
        revisedEnd: 3,
      },
    ],
    sourceRuns: [{ startOffset: 0, endOffset: 3, formatting: { bold: true } }],
    targetRuns: [{ startOffset: 0, endOffset: 3, formatting: { italic: true } }],
    authoredChanges: [],
  },
});

describe("DocxComparisonProgram", () => {
  test("captures and deeply freezes the sole instruction payload", () => {
    const input = replacement();
    const program = DocxComparisonProgram.create([input]);

    Reflect.set(input.range.segments[0]!, "text", "corrupted");
    Reflect.set(input.range.sourceRuns[0]!.formatting, "bold", false);

    const [instruction] = program.consume();
    expect(instruction?.type).toBe("replaceText");
    if (instruction?.type !== "replaceText") throw new Error("expected replacement");
    expect(instruction.range.sourceText).toBe("old");
    expect(instruction.range.fragments[0]?.text).toBe("old");
    expect(
      instruction.range.fragments.find(({ type }) => type === "del")?.sourceFormatting.bold,
    ).toBe(true);
    expect(Object.isFrozen(instruction)).toBe(true);
    expect(Object.isFrozen(instruction.range.fragments)).toBe(true);
    expect(Object.isFrozen(instruction.range.fragments[0])).toBe(true);
  });

  test("is consumable exactly once", () => {
    const program = DocxComparisonProgram.create([replacement()]);
    expect(program.consume()).toHaveLength(1);
    expect(() => program.consume()).toThrow("consumed more than once");
  });

  test("rejects a range whose canonical text does not name its source", () => {
    const input = replacement();
    Reflect.set(input.range, "sourceText", "elsewhere");
    expect(() => DocxComparisonProgram.create([input])).toThrow(
      "does not name its exact source and target text",
    );
  });

  test("rejects an authored formatting delta missing from canonical changes", () => {
    const input = replacement();
    Reflect.set(input.range, "segments", [
      {
        type: "equal",
        text: "old",
        baseStart: 0,
        baseEnd: 3,
        revisedStart: 0,
        revisedEnd: 3,
      },
    ]);
    Reflect.set(input.range, "targetText", "old");
    expect(() => DocxComparisonProgram.create([input])).toThrow(
      "Canonical authored-formatting changes do not reconstruct the target run",
    );
  });

  test("rejects oversized instruction and geometry arrays before copying them", () => {
    expect(() =>
      DocxComparisonProgram.create(Array.from({ length: 10_001 }, replacement)),
    ).toThrow("exceeds its instruction limit");

    const pairing = {
      base: { tableIndex: 0, rowIndex: 0, cellIndex: 0 },
      target: { tableIndex: 0, rowIndex: 0, cellIndex: 0 },
    };
    expect(() =>
      DocxComparisonProgram.create([
        {
          type: "matchTableGeometry",
          pairings: Array.from({ length: 10_001 }, () => pairing),
        },
      ]),
    ).toThrow("exceeds its pairing limit");
  });
});
