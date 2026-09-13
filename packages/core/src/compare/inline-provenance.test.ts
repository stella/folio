import { describe, expect, test } from "bun:test";

import type { Paragraph, TextFormatting } from "../types/document";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import { createEmptyDocument } from "../utils/createDocument";
import { FolioDocxReviewer } from "../ai-edits/headless";
import { compareDocx } from "./compare";

const OPTIONS = { author: "compare", timestamp: "2026-09-13T00:00:00.000Z" } as const;

type RunSpec = { text: string; formatting?: TextFormatting };

const MIXED_TARGET_RUNS = [
  {
    text: "Revised ",
    formatting: {
      bold: true,
      boldCs: false,
      italic: true,
      italicCs: false,
      fontFamily: {
        ascii: "Aptos",
        hAnsi: "Aptos",
        eastAsia: "Yu Gothic",
        cs: "Noto Naskh Arabic",
      },
      fontSize: 25,
      fontSizeCs: 27,
      color: { themeColor: "accent1", themeTint: "80" },
      underline: { style: "double", color: { rgb: "C00000" } },
      highlight: "yellow",
    },
  },
  {
    text: "terms",
    formatting: {
      bold: false,
      boldCs: true,
      italic: false,
      italicCs: true,
      fontFamily: { ascii: "Courier New", cs: "Arial" },
      fontSize: 22,
      fontSizeCs: 24,
      color: { auto: true },
      underline: { style: "single", color: { themeColor: "accent2" } },
      highlight: "cyan",
    },
  },
] as const satisfies readonly RunSpec[];

const paragraph = (paraId: string, runs: readonly RunSpec[]): Paragraph => ({
  type: "paragraph",
  paraId,
  textId: paraId,
  content: runs.map(({ text, formatting }) => ({
    type: "run",
    ...(formatting !== undefined && { formatting }),
    content: [{ type: "text", text }],
  })),
});

const documentWith = (paragraphs: readonly Paragraph[]): Promise<ArrayBuffer> => {
  const document = createEmptyDocument();
  document.package.document.content = [...paragraphs];
  return createDocx(document);
};

type FormattedCharacter = {
  paragraphIndex: number;
  text: string;
  formatting: TextFormatting | undefined;
};

const directFormattingCharacters = async (
  buffer: ArrayBuffer,
): Promise<readonly FormattedCharacter[]> => {
  const document = await parseDocx(buffer, { detectVariables: false, preloadFonts: false });
  const characters: FormattedCharacter[] = [];
  for (const [paragraphIndex, block] of document.package.document.content.entries()) {
    if (block.type !== "paragraph") continue;
    for (const run of block.content) {
      if (run.type !== "run") continue;
      for (const content of run.content) {
        if (content.type !== "text") continue;
        for (const text of Array.from(content.text)) {
          characters.push({ paragraphIndex, text, formatting: run.formatting });
        }
      }
    }
  }
  return characters;
};

const expectAcceptedAndRejectedFormatting = async ({
  base,
  revised,
}: {
  base: ArrayBuffer;
  revised: ArrayBuffer;
}): Promise<void> => {
  const [expectedBase, expectedRevised] = await Promise.all([
    directFormattingCharacters(base),
    directFormattingCharacters(revised),
  ]);
  const compared = await compareDocx(base, revised, OPTIONS);
  if (compared.isErr()) throw compared.error;

  const accepting = await FolioDocxReviewer.fromBuffer(compared.value.buffer);
  expect(accepting.acceptAll()).toBeGreaterThan(0);
  expect(await directFormattingCharacters(await accepting.toBuffer())).toEqual(expectedRevised);

  const rejecting = await FolioDocxReviewer.fromBuffer(compared.value.buffer);
  expect(rejecting.rejectAll()).toBeGreaterThan(0);
  expect(await directFormattingCharacters(await rejecting.toBuffer())).toEqual(expectedBase);
};

describe("inline formatting provenance through comparison", () => {
  test("accepting a replacement retains every revised direct run property", async () => {
    await expectAcceptedAndRejectedFormatting({
      base: await documentWith([paragraph("10000001", [{ text: "Base terms" }])]),
      revised: await documentWith([paragraph("20000001", MIXED_TARGET_RUNS)]),
    });
  });

  test("accepting an inserted paragraph retains target direct run properties", async () => {
    await expectAcceptedAndRejectedFormatting({
      base: await documentWith([paragraph("30000001", [{ text: "Anchor" }])]),
      revised: await documentWith([
        paragraph("40000001", MIXED_TARGET_RUNS),
        paragraph("30000001", [{ text: "Anchor" }]),
      ]),
    });
  });

  test("a text-equal direct run-property change has distinct accepted and rejected values", async () => {
    await expectAcceptedAndRejectedFormatting({
      base: await documentWith([
        paragraph("50000001", [
          {
            text: "Stable wording",
            formatting: {
              bold: false,
              boldCs: false,
              italic: false,
              italicCs: false,
              color: { auto: true },
              underline: { style: "none" },
            },
          },
        ]),
      ]),
      revised: await documentWith([
        paragraph("50000001", [
          {
            text: "Stable wording",
            formatting: {
              bold: true,
              boldCs: true,
              italic: true,
              italicCs: true,
              color: { themeColor: "accent3", themeShade: "40" },
              underline: { style: "dash", color: { rgb: "0070C0" } },
              highlight: "green",
            },
          },
        ]),
      ]),
    });
  });
});
