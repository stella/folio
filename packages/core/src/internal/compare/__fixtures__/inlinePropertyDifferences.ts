// A field-totality gate in a test file is read by no compiler, so the
// difference cases per inline formatting property live here, where the
// package typecheck compiles them.
import type {
  FolioContentInlineFormatting,
  FolioContentInlineFormattingPatch,
  FolioContentRun,
} from "../../../compare/content-types";

/** A run's presentation: everything about it but its text. */
export type RunPresentation = Omit<FolioContentRun, "text">;

type DifferenceFixture = Readonly<{
  label: string;
  base: RunPresentation;
  revised: RunPresentation;
  expected: FolioContentInlineFormattingPatch;
}>;

export const PROPERTY_DIFFERENCES = {
  bold: [
    {
      label: "effective on",
      base: {},
      revised: { bold: true },
      expected: { bold: true },
    },
    {
      label: "authored on",
      base: { bold: true },
      revised: { bold: true, directFormatting: { bold: true } },
      expected: { bold: true },
    },
    {
      label: "authored removal",
      base: { bold: true, directFormatting: { bold: true } },
      revised: { bold: true },
      expected: { bold: null },
    },
  ],
  italic: [
    {
      label: "effective on",
      base: {},
      revised: { italic: true },
      expected: { italic: true },
    },
    {
      label: "authored on",
      base: { italic: true },
      revised: { italic: true, directFormatting: { italic: true } },
      expected: { italic: true },
    },
    {
      label: "authored removal",
      base: { italic: true, directFormatting: { italic: true } },
      revised: { italic: true },
      expected: { italic: null },
    },
  ],
  underline: [
    {
      label: "effective on",
      base: {},
      revised: { underline: true },
      expected: { underline: true },
    },
    {
      label: "authored on",
      base: { underline: true },
      revised: { underline: true, directFormatting: { underline: true } },
      expected: { underline: true },
    },
    {
      label: "authored removal",
      base: { underline: true, directFormatting: { underline: true } },
      revised: { underline: true },
      expected: { underline: null },
    },
  ],
  strike: [
    {
      label: "effective on",
      base: {},
      revised: { strike: true },
      expected: { strike: true },
    },
    {
      label: "authored on",
      base: { strike: true },
      revised: { strike: true, directFormatting: { strike: true } },
      expected: { strike: true },
    },
    {
      label: "authored removal",
      base: { strike: true, directFormatting: { strike: true } },
      revised: { strike: true },
      expected: { strike: null },
    },
  ],
  fontFamily: [
    {
      label: "effective value",
      base: {},
      revised: { fontFamily: "Georgia" },
      expected: { fontFamily: "Georgia" },
    },
    {
      label: "authored value",
      base: { fontFamily: "Georgia" },
      revised: {
        fontFamily: "Georgia",
        directFormatting: { fontFamily: "Georgia" },
      },
      expected: { fontFamily: "Georgia" },
    },
    {
      label: "authored removal",
      base: {
        fontFamily: "Georgia",
        directFormatting: { fontFamily: "Georgia" },
      },
      revised: { fontFamily: "Georgia" },
      expected: { fontFamily: null },
    },
  ],
  fontSizePt: [
    {
      label: "effective value",
      base: {},
      revised: { fontSizePt: 10.5 },
      expected: { fontSizePt: 10.5 },
    },
    {
      label: "authored value",
      base: { fontSizePt: 10.5 },
      revised: { fontSizePt: 10.5, directFormatting: { fontSizePt: 10.5 } },
      expected: { fontSizePt: 10.5 },
    },
    {
      label: "authored removal",
      base: { fontSizePt: 10.5, directFormatting: { fontSizePt: 10.5 } },
      revised: { fontSizePt: 10.5 },
      expected: { fontSizePt: null },
    },
  ],
  color: [
    {
      label: "effective value",
      base: {},
      revised: { color: "#c00000" },
      expected: { color: "C00000" },
    },
    {
      label: "authored value",
      base: { color: "C00000" },
      revised: { color: "c00000", directFormatting: { color: "#c00000" } },
      expected: { color: "C00000" },
    },
    {
      label: "authored removal",
      base: { color: "C00000", directFormatting: { color: "c00000" } },
      revised: { color: "#C00000" },
      expected: { color: null },
    },
  ],
} as const satisfies Record<
  keyof FolioContentInlineFormatting,
  readonly DifferenceFixture[]
>;
