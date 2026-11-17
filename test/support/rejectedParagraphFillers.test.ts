import { panic } from "better-result";
import { expect, test } from "bun:test";

import type {
  BlockContent,
  Document,
  Paragraph,
} from "../../packages/docx-core/src/model/document";
import { rejectedParagraphFillerIdentities } from "./rejectedParagraphFillers";

const paragraph = (paraId: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [],
});
const outside = {
  ...paragraph("10000000"),
  sectionProperties: { pageWidth: 11906 },
};
const control = (content: BlockContent[]) =>
  ({
    type: "blockSdt",
    properties: { sdtType: "group", id: 9 },
    content,
  }) as const;
const custom = (content: BlockContent[]) =>
  ({
    type: "blockCustomXml",
    openingXml: '<w:customXml w:element="clause">',
    closingXml: "</w:customXml>",
    content,
  }) as const;
const cell = (content: BlockContent[]) =>
  ({
    type: "table",
    rows: [{ type: "tableRow", cells: [{ type: "tableCell", content }] }],
  }) as const;
const owners = {
  control,
  custom,
  cell,
  nested: (content: BlockContent[]) => control([custom(content)]),
  "wrapped-cell": (content: BlockContent[]) => cell([control(content)]),
} as const;
const document = (content: BlockContent[]): Document => ({
  package: { document: { content } },
});

type CompareOptions = Parameters<typeof rejectedParagraphFillerIdentities>[0];
/** These literal test documents contain only JSON fields. Preserve every field except proven IDs. */
const compared = ({ source, actual, expected }: CompareOptions) => {
  const ids = new Map(
    rejectedParagraphFillerIdentities({ source, actual, expected }).map(
      ({ actualId, expectedId }) => [actualId, expectedId],
    ),
  );
  return JSON.parse(
    JSON.stringify(actual, (key, value) =>
      key === "paraId" && typeof value === "string"
        ? (ids.get(value) ?? value)
        : value,
    ),
  );
};

for (const [name, owner] of Object.entries(owners))
  test(`${name}: only an exhausted bare added-mark container permits filler identity equivalence`, () => {
    for (const kind of ["ins", "moveTo"] as const)
      for (const length of [2, 3, 4]) {
        const added = Array.from({ length }, (_, index) => ({
          ...paragraph((index + 1).toString(16).padStart(8, "0")),
          pPrMark: { kind, info: { id: index + 1, author: "Reviewer" } },
        }));
        const source = document([owner(added), outside]);
        const actual = document([
          owner([paragraph(length.toString(16).padStart(8, "0"))]),
          outside,
        ]);
        const expected = document([owner([paragraph("00000001")]), outside]);
        expect(
          rejectedParagraphFillerIdentities({ source, actual, expected }),
        ).toEqual([
          {
            actualId: length.toString(16).padStart(8, "0"),
            expectedId: "00000001",
          },
        ]);
        expect(compared({ source, actual, expected })).toEqual(expected);
      }
  });

const added = ["00000001", "00000002"].map(
  (id, index) =>
    Object.assign(paragraph(id), {
      pPrMark: { kind: "ins", info: { id: index + 1, author: "Reviewer" } },
    }) satisfies Paragraph,
);
const firstAdded = added.at(0) ?? panic("Two source paragraphs exist.");
const secondAdded = added.at(1) ?? panic("Two source paragraphs exist.");
const source = document([control(added), outside]);
const expected = document([control([paragraph("00000001")]), outside]);
const actual = document([control([paragraph("00000002")]), outside]);

test("stable empty paragraph IDs remain exact outside and inside the candidate", () => {
  const drift = document([
    control([paragraph("00000002")]),
    { ...outside, paraId: "10000001" },
  ]);
  expect(compared({ source, actual: drift, expected })).not.toEqual(expected);
  const mixed = document([
    control([paragraph("00000001"), secondAdded]),
    outside,
  ]);
  expect(
    rejectedParagraphFillerIdentities({ source: mixed, actual, expected }),
  ).toEqual([]);
  expect(compared({ source: mixed, actual, expected })).not.toEqual(expected);
});

test("missing, extra, authored or unlocated output paragraphs cannot be normalized", () => {
  const variants = {
    missing: document([control([]), outside]),
    extra: document([
      control([paragraph("00000002"), paragraph("00000004")]),
      outside,
    ]),
    text: document([
      control([
        {
          ...paragraph("00000002"),
          content: [
            { type: "run", content: [{ type: "text", text: "Lost words" }] },
          ],
        },
      ]),
      outside,
    ]),
    section: document([
      control([
        { ...paragraph("00000002"), sectionProperties: { pageWidth: 12240 } },
      ]),
      outside,
    ]),
    range: document([
      control([
        {
          ...paragraph("00000002"),
          content: [{ type: "bookmarkStart", id: 6, name: "anchor" }],
        },
      ]),
      outside,
    ]),
    formatting: document([
      control([
        { ...paragraph("00000002"), formatting: { alignment: "center" } },
      ]),
      outside,
    ]),
    owner: document([
      {
        ...control([paragraph("00000002")]),
        properties: { sdtType: "group", id: 10 },
      },
      outside,
    ]),
    moved: document([outside, control([paragraph("00000002")])]),
    foreign: document([control([paragraph("00000004")]), outside]),
  } satisfies Record<string, Document>;
  for (const changed of Object.values(variants)) {
    expect(
      rejectedParagraphFillerIdentities({ source, actual: changed, expected }),
    ).toEqual([]);
    expect(compared({ source, actual: changed, expected })).not.toEqual(
      expected,
    );
  }
});

test("source sections, ranges, formatting, unknown fields and nonparagraph siblings exclude equivalence", () => {
  const variants = [
    [{ ...firstAdded, sectionProperties: { pageWidth: 11906 } }, secondAdded],
    [{ ...firstAdded, formatting: { alignment: "center" } }, secondAdded],
    [{ ...firstAdded, content: [{ type: "bookmarkEnd", id: 6 }] }, secondAdded],
    [
      {
        ...firstAdded,
        pPrMark: { ...firstAdded.pPrMark, preservedXml: "<w:unknown/>" },
      },
      secondAdded,
    ],
    [firstAdded, { type: "bookmarkEnd", id: 6 }, secondAdded],
  ] satisfies BlockContent[][];
  for (const content of variants) {
    const original = document([control(content), outside]);
    expect(
      rejectedParagraphFillerIdentities({ source: original, actual, expected }),
    ).toEqual([]);
    expect(compared({ source: original, actual, expected })).not.toEqual(
      expected,
    );
  }
});

test("outside paragraph loss, text, section and range drift remain observable", () => {
  const variants = [
    document([control([paragraph("00000002")])]),
    document([
      control([paragraph("00000002")]),
      { ...outside, sectionProperties: { pageWidth: 12240 } },
    ]),
    document([
      control([paragraph("00000002")]),
      {
        ...outside,
        content: [{ type: "run", content: [{ type: "text", text: "Drift" }] }],
      },
    ]),
    document([
      control([paragraph("00000002")]),
      { ...outside, content: [{ type: "bookmarkEnd", id: 6 }] },
    ]),
  ];
  for (const changed of variants)
    expect(compared({ source, actual: changed, expected })).not.toEqual(
      expected,
    );
});
