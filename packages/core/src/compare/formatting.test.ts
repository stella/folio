import { describe, expect, test } from "bun:test";

import type {
  FolioContentBlock,
  FolioContentPropertyInputValue,
  FolioContentPropertySet,
} from "./content-types";
import { contentBlockFixture } from "./content-test-fixtures";
import { inlineFormattingSegments } from "./formatting";

type TestInlineFormatting = {
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly underline?: boolean;
  readonly strike?: boolean;
  readonly fontFamily?: string;
  readonly fontSizePt?: number;
  readonly color?: string;
};

type TestRun = TestInlineFormatting & {
  readonly text: string;
  readonly directFormatting?: TestInlineFormatting;
};

const propertySet = (formatting: TestInlineFormatting): FolioContentPropertySet =>
  Object.freeze(
    Object.entries(formatting)
      .filter((entry): entry is [string, FolioContentPropertyInputValue] => entry[1] !== undefined)
      .map(([key, value]) => Object.freeze({ key, value }))
      .toSorted((left, right) => {
        if (left.key < right.key) return -1;
        if (left.key > right.key) return 1;
        return 0;
      }),
  );

const block = (runs: readonly TestRun[]): FolioContentBlock =>
  Object.freeze({
    ...contentBlockFixture("block", "Contract"),
    runs: Object.freeze(
      runs.map(({ text, directFormatting, ...effectiveFormatting }) =>
        Object.freeze({
          text,
          effectiveFormatting: propertySet(effectiveFormatting),
          authoredFormatting: propertySet(directFormatting ?? {}),
        }),
      ),
    ),
  });

const DIRECT_BOOLEAN_STATES = [
  { label: "absent" },
  { label: "on", value: true },
  { label: "off", value: false },
] as const;

type DirectBooleanState = (typeof DIRECT_BOOLEAN_STATES)[number];

const directBooleanPresence = (state: DirectBooleanState) => {
  if ("value" in state) return { type: "present", value: state.value } as const;
  return { type: "absent" } as const;
};

const effectiveBooleanPresence = (value: boolean) => {
  if (value) return { type: "present", value: true } as const;
  return { type: "absent" } as const;
};

const INLINE_BOOLEAN_PROPERTIES = [
  "bold",
  "italic",
  "underline",
  "strike",
] as const satisfies readonly (keyof Pick<
  TestInlineFormatting,
  "bold" | "italic" | "underline" | "strike"
>)[];

const blockWithBooleanState = (
  property: (typeof INLINE_BOOLEAN_PROPERTIES)[number],
  inherited: boolean,
  direct: DirectBooleanState,
): FolioContentBlock => {
  const effective = "value" in direct ? direct.value : inherited;
  return block([
    {
      text: "Contract",
      ...(effective && { [property]: true }),
      ...("value" in direct && { directFormatting: { [property]: direct.value } }),
    },
  ]);
};

describe("inlineFormattingSegments", () => {
  test("treats equivalent formatting across different run splits as equal", () => {
    const split = block([
      { text: "", bold: true },
      { text: "Con", strike: true },
      { text: "tract", strike: true },
      { text: "", italic: true },
    ]);
    const joined = block([{ text: "Contract", strike: true }]);

    expect(
      inlineFormattingSegments({ baseBlock: split, targetBlock: joined, maxSegments: 10 }),
    ).toEqual([]);
  });

  test("coalesces adjacent strike changes across target run boundaries", () => {
    expect(
      inlineFormattingSegments({
        baseBlock: block([{ text: "Contract" }]),
        targetBlock: block([
          { text: "Con", strike: true },
          { text: "tract", strike: true },
        ]),
        maxSegments: 10,
      }),
    ).toEqual([
      {
        startOffset: 0,
        endOffset: 8,
        formatting: {
          authored: [],
          effective: [
            {
              key: "strike",
              base: { type: "absent" },
              revised: { type: "present", value: true },
            },
          ],
        },
      },
    ]);
  });

  test("reports canonical target font face, half-point size, and RGB color", () => {
    expect(
      inlineFormattingSegments({
        baseBlock: block([{ text: "Contract", fontFamily: "Arial", fontSizePt: 10 }]),
        targetBlock: block([
          {
            text: "Contract",
            fontFamily: "Georgia",
            fontSizePt: 10.5,
            color: "C00000",
            directFormatting: {
              fontFamily: "Georgia",
              fontSizePt: 10.5,
              color: "C00000",
            },
          },
        ]),
        maxSegments: 10,
      }),
    ).toEqual([
      {
        startOffset: 0,
        endOffset: 8,
        formatting: {
          authored: [
            {
              key: "color",
              base: { type: "absent" },
              revised: { type: "present", value: "C00000" },
            },
            {
              key: "fontFamily",
              base: { type: "absent" },
              revised: { type: "present", value: "Georgia" },
            },
            {
              key: "fontSizePt",
              base: { type: "absent" },
              revised: { type: "present", value: 10.5 },
            },
          ],
          effective: [
            {
              key: "color",
              base: { type: "absent" },
              revised: { type: "present", value: "C00000" },
            },
            {
              key: "fontFamily",
              base: { type: "present", value: "Arial" },
              revised: { type: "present", value: "Georgia" },
            },
            {
              key: "fontSizePt",
              base: { type: "present", value: 10 },
              revised: { type: "present", value: 10.5 },
            },
          ],
        },
      },
    ]);
  });

  test("represents cleared direct font properties explicitly", () => {
    expect(
      inlineFormattingSegments({
        baseBlock: block([
          {
            text: "Contract",
            fontFamily: "Georgia",
            fontSizePt: 10.5,
            color: "C00000",
            directFormatting: {
              fontFamily: "Georgia",
              fontSizePt: 10.5,
              color: "C00000",
            },
          },
        ]),
        targetBlock: block([{ text: "Contract" }]),
        maxSegments: 10,
      }),
    ).toEqual([
      {
        startOffset: 0,
        endOffset: 8,
        formatting: {
          authored: [
            {
              key: "color",
              base: { type: "present", value: "C00000" },
              revised: { type: "absent" },
            },
            {
              key: "fontFamily",
              base: { type: "present", value: "Georgia" },
              revised: { type: "absent" },
            },
            {
              key: "fontSizePt",
              base: { type: "present", value: 10.5 },
              revised: { type: "absent" },
            },
          ],
          effective: [
            {
              key: "color",
              base: { type: "present", value: "C00000" },
              revised: { type: "absent" },
            },
            {
              key: "fontFamily",
              base: { type: "present", value: "Georgia" },
              revised: { type: "absent" },
            },
            {
              key: "fontSizePt",
              base: { type: "present", value: 10.5 },
              revised: { type: "absent" },
            },
          ],
        },
      },
    ]);
  });

  test("verification distinguishes an inherited value from the same direct value", () => {
    const inherited = block([{ text: "Contract", fontFamily: "Georgia" }]);
    const direct = block([
      {
        text: "Contract",
        fontFamily: "Georgia",
        directFormatting: { fontFamily: "Georgia" },
      },
    ]);

    expect(
      inlineFormattingSegments({ baseBlock: inherited, targetBlock: direct, maxSegments: 10 }),
    ).toEqual([
      {
        startOffset: 0,
        endOffset: 8,
        formatting: {
          authored: [
            {
              key: "fontFamily",
              base: { type: "absent" },
              revised: { type: "present", value: "Georgia" },
            },
          ],
          effective: [],
        },
      },
    ]);
  });

  test("preserves direct boolean on, off, and absence across every inherited state", () => {
    for (const property of INLINE_BOOLEAN_PROPERTIES) {
      for (const inherited of [false, true]) {
        for (const baseDirect of DIRECT_BOOLEAN_STATES) {
          for (const targetDirect of DIRECT_BOOLEAN_STATES) {
            const base = blockWithBooleanState(property, inherited, baseDirect);
            const target = blockWithBooleanState(property, inherited, targetDirect);
            const sameDirectState = baseDirect.label === targetDirect.label;
            const baseEffective = "value" in baseDirect ? baseDirect.value : inherited;
            const targetEffective = "value" in targetDirect ? targetDirect.value : inherited;
            const expectedSegments = sameDirectState
              ? []
              : [
                  {
                    startOffset: 0,
                    endOffset: 8,
                    formatting: {
                      authored: [
                        {
                          key: property,
                          base: directBooleanPresence(baseDirect),
                          revised: directBooleanPresence(targetDirect),
                        },
                      ],
                      effective:
                        baseEffective === targetEffective
                          ? []
                          : [
                              {
                                key: property,
                                base: effectiveBooleanPresence(baseEffective),
                                revised: effectiveBooleanPresence(targetEffective),
                              },
                            ],
                    },
                  },
                ];

            expect({
              property,
              inherited,
              base: baseDirect.label,
              target: targetDirect.label,
              segments: inlineFormattingSegments({
                baseBlock: base,
                targetBlock: target,
                maxSegments: 10,
              }),
            }).toEqual({
              property,
              inherited,
              base: baseDirect.label,
              target: targetDirect.label,
              segments: expectedSegments,
            });
          }
        }
      }
    }
  });
});
