import { expect, test } from "bun:test";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { createDocx } from "../docx/rezip";
import { createEmptyDocument } from "../utils/createDocument";
import { compareDocx } from "./compare";

const paragraph = (text: string, paraId: string, numId?: number) => ({
  type: "paragraph" as const,
  paraId,
  textId: paraId,
  ...(numId === undefined ? {} : { formatting: { numPr: { numId, ilvl: 0 } } }),
  content: [{ type: "run" as const, content: [{ type: "text" as const, text }] }],
});

const document = (introduced: boolean) => {
  const result = createEmptyDocument();
  result.package.document.content = [
    paragraph("Anchor paragraph.", "11111111"),
    ...(introduced ? [paragraph("Introduced list item.", "22222222", 5)] : []),
  ];
  if (introduced) {
    result.package.numbering = {
      abstractNums: [
        { abstractNumId: 5, levels: [{ ilvl: 0, numFmt: "decimal", lvlText: "%1." }] },
      ],
      nums: [{ numId: 5, abstractNumId: 5 }],
    };
  }
  return createDocx(result);
};

const collisionDocument = ({
  markerBold = false,
  numFmt,
}: {
  markerBold?: boolean;
  numFmt: "decimal" | "lowerRoman";
}) => {
  const result = createEmptyDocument();
  result.package.document.content = [paragraph("Shared list item.", "33333333", 5)];
  result.package.numbering = {
    abstractNums: [
      {
        abstractNumId: 5,
        levels: [
          { ilvl: 0, numFmt, lvlText: "%1.", ...(markerBold ? { rPr: { bold: true } } : {}) },
        ],
      },
    ],
    nums: [{ numId: 5, abstractNumId: 5 }],
  };
  return createDocx(result);
};

const collisionTableDocument = (
  withInsertedRow: boolean,
  numFmt: "decimal" | "lowerRoman",
) => {
  const result = createEmptyDocument();
  result.package.document.content = [
    {
      type: "table",
      rows: [
        {
          type: "tableRow",
          cells: [
            {
              type: "tableCell",
              content: [paragraph("Shared table list item.", "44444444", 5)],
            },
          ],
        },
        ...(withInsertedRow
          ? [
              {
                type: "tableRow" as const,
                cells: [
                  {
                    type: "tableCell" as const,
                    content: [paragraph("Inserted table list item.", "55555555", 5)],
                  },
                ],
              },
            ]
          : []),
      ],
    },
  ];
  result.package.numbering = {
    abstractNums: [
      { abstractNumId: 5, levels: [{ ilvl: 0, numFmt, lvlText: "%1." }] },
    ],
    nums: [{ numId: 5, abstractNumId: 5 }],
  };
  return createDocx(result);
};

test("imports target-only numbering for an introduced list and preserves rejection", async () => {
  const result = await compareDocx(await document(false), await document(true), {
    author: "compare",
    timestamp: "2026-09-13T00:00:00.000Z",
  });
  if (result.isErr()) throw result.error;
  expect(result.value.verification).toEqual({ status: "verified" });

  const accepting = await FolioDocxReviewer.fromBuffer(result.value.buffer);
  expect(accepting.acceptAll()).toBeGreaterThan(0);
  expect(accepting.snapshot().blocks.at(1)?.listReference).toEqual({ numId: 5, level: 0 });

  const rejecting = await FolioDocxReviewer.fromBuffer(result.value.buffer);
  expect(rejecting.rejectAll()).toBeGreaterThan(0);
  expect(rejecting.snapshot().blocks.map(({ text }) => text)).toEqual(["Anchor paragraph."]);
});

test("rebinds a colliding target numbering definition through a tracked paragraph change", async () => {
  const result = await compareDocx(
    await collisionDocument({ numFmt: "decimal" }),
    await collisionDocument({ numFmt: "lowerRoman" }),
    {
      author: "compare",
      timestamp: "2026-09-13T00:00:00.000Z",
    },
  );
  if (result.isErr()) throw result.error;
  expect(result.value.verification).toEqual({ status: "verified" });

  const accepting = await FolioDocxReviewer.fromBuffer(result.value.buffer);
  expect(accepting.acceptAll()).toBeGreaterThan(0);
  expect(accepting.snapshot().blocks.at(0)?.listReference).toEqual({ numId: 6, level: 0 });
  expect(
    accepting.readNumberingDefinitions().find(({ numId, level }) => numId === 6 && level === 0)
      ?.format,
  ).toBe("lowerRoman");

  const rejecting = await FolioDocxReviewer.fromBuffer(result.value.buffer);
  expect(rejecting.rejectAll()).toBeGreaterThan(0);
  expect(rejecting.snapshot().blocks.at(0)?.listReference).toEqual({ numId: 5, level: 0 });
  expect(
    rejecting.readNumberingDefinitions().find(({ numId, level }) => numId === 5 && level === 0)
      ?.format,
  ).toBe("decimal");
});

test("rebinds a colliding marker-typography change even when its label is unchanged", async () => {
  const result = await compareDocx(
    await collisionDocument({ numFmt: "decimal" }),
    await collisionDocument({ numFmt: "decimal", markerBold: true }),
    {
      author: "compare",
      timestamp: "2026-09-13T00:00:00.000Z",
    },
  );
  if (result.isErr()) throw result.error;
  expect(result.value.verification).toEqual({ status: "verified" });

  const accepting = await FolioDocxReviewer.fromBuffer(result.value.buffer);
  expect(accepting.acceptAll()).toBeGreaterThan(0);
  const acceptedNumbering = accepting.toDocument().package.numbering;
  const acceptedNum = acceptedNumbering?.nums.find(({ numId }) => numId === 6);
  expect(
    acceptedNumbering?.abstractNums
      .find(({ abstractNumId }) => abstractNumId === acceptedNum?.abstractNumId)
      ?.levels.find(({ ilvl }) => ilvl === 0)?.rPr?.bold,
  ).toBe(true);

  const rejecting = await FolioDocxReviewer.fromBuffer(result.value.buffer);
  expect(rejecting.rejectAll()).toBeGreaterThan(0);
  const rejectedNumbering = rejecting.toDocument().package.numbering;
  const rejectedNum = rejectedNumbering?.nums.find(({ numId }) => numId === 5);
  expect(
    rejectedNumbering?.abstractNums
      .find(({ abstractNumId }) => abstractNumId === rejectedNum?.abstractNumId)
      ?.levels.find(({ ilvl }) => ilvl === 0)?.rPr?.bold,
  ).toBeUndefined();
});

test("rebinds a colliding numbering definition in an inserted table row", async () => {
  const result = await compareDocx(
    await collisionTableDocument(false, "decimal"),
    await collisionTableDocument(true, "lowerRoman"),
    {
      author: "compare",
      timestamp: "2026-09-13T00:00:00.000Z",
    },
  );
  if (result.isErr()) throw result.error;
  expect(result.value.verification).toEqual({ status: "verified" });

  const accepting = await FolioDocxReviewer.fromBuffer(result.value.buffer);
  expect(accepting.acceptAll()).toBeGreaterThan(0);
  const accepted = await FolioDocxReviewer.fromBuffer(await accepting.toBuffer());
  expect(accepted.snapshot().blocks.map(({ text, listReference }) => ({ text, listReference }))).toEqual([
    { text: "Shared table list item.", listReference: { numId: 6, level: 0 } },
    { text: "Inserted table list item.", listReference: { numId: 6, level: 0 } },
  ]);
  expect(
    accepted.readNumberingDefinitions().find(({ numId, level }) => numId === 6 && level === 0)
      ?.format,
  ).toBe("lowerRoman");

  const rejecting = await FolioDocxReviewer.fromBuffer(result.value.buffer);
  expect(rejecting.rejectAll()).toBeGreaterThan(0);
  const rejected = await FolioDocxReviewer.fromBuffer(await rejecting.toBuffer());
  expect(rejected.snapshot().blocks.map(({ text, listReference }) => ({ text, listReference }))).toEqual([
    { text: "Shared table list item.", listReference: { numId: 5, level: 0 } },
  ]);
});
