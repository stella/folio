import { expect, setDefaultTimeout, spyOn, test } from "bun:test";
import { panic } from "better-result";
import fc from "fast-check";

import { assertProperty, propertyTestTimeout } from "../../../test/property-testing";
import { expectParagraphBlock } from "../../../test/paragraphBlock";
import { FolioDocxReviewer } from "./ai-edits/headless";
import { createDocx } from "./docx/rezip";
import { generateRedlineDocx, GenerateRedlineDocxResourceImportError } from "./redline";
import { createEmptyDocument } from "./utils/createDocument";
import { alignFolioBlocks } from "./version-comparison";

setDefaultTimeout(propertyTestTimeout(30_000));

const UNSUPPORTED_STYLE = "UnsupportedStyle";
const SUPPORTED_STYLE = "SupportedStyle";
const NORMAL_STYLE = "Normal";
const INSERTION_CASES = [false, true].flatMap((removal) =>
  [false, true].map((independent) => ({ removal, independent })),
);

const paragraph = (paraId: string, text: string, styleId = NORMAL_STYLE) => ({
  type: "paragraph" as const,
  paraId,
  textId: paraId,
  formatting: { styleId },
  content: [{ type: "run" as const, content: [{ type: "text" as const, text }] }],
});

type ResourceFixtureOptions = {
  side: "base" | "revised";
  removal: boolean;
  independent: boolean;
  suffix: string;
  styleCase: "missingLinkedStyle" | "unknownDirectStyle";
};

const resourceFixture = ({
  side,
  removal,
  independent,
  suffix,
  styleCase,
}: ResourceFixtureOptions) => {
  const document = createEmptyDocument();
  document.package.styles = {
    styles: [
      { type: "paragraph", styleId: NORMAL_STYLE, name: "Normal", default: true },
      ...(styleCase === "missingLinkedStyle"
        ? [
            {
              type: "paragraph",
              styleId: UNSUPPORTED_STYLE,
              name: "Unsupported style",
              ...(side === "revised" && { link: "MissingStyle" }),
              rPr: { bold: side === "revised" },
            } as const,
          ]
        : []),
      {
        type: "paragraph",
        styleId: SUPPORTED_STYLE,
        name: "Supported style",
        rPr: { italic: side === "revised" },
      },
    ],
  };
  document.package.document.content = [
    paragraph("91000001", `Leading anchor ${suffix}`),
    ...(side === "base" && removal ? [paragraph("91000002", `Removed original ${suffix}`)] : []),
    paragraph("91000003", `Middle anchor ${suffix}`),
    ...(side === "revised"
      ? [paragraph("92000004", `Unsupported replacement ${suffix}`, UNSUPPORTED_STYLE)]
      : []),
    paragraph("91000005", `Trailing anchor ${suffix}`),
    ...(side === "revised" && independent
      ? [paragraph("92000006", `Independent insertion ${suffix}`, SUPPORTED_STYLE)]
      : []),
  ];
  return createDocx(document);
};

test("a failing insertion style closure refuses the whole redline without changing input packages", async () => {
  const applySpy = spyOn(FolioDocxReviewer.prototype, "applyDocumentOperationsToStory");
  try {
    await assertProperty(
      fc.asyncProperty(fc.stringMatching(/^[a-z]{1,20}$/u), async (suffix) => {
        for (const styleCase of ["missingLinkedStyle", "unknownDirectStyle"] as const) {
          for (const { removal, independent } of INSERTION_CASES) {
            applySpy.mockClear();
            const base = await resourceFixture({
              side: "base",
              removal,
              independent,
              suffix,
              styleCase,
            });
            const revised = await resourceFixture({
              side: "revised",
              removal,
              independent,
              suffix,
              styleCase,
            });
            const baseBytes = new Uint8Array(base).slice();
            const revisedBytes = new Uint8Array(revised).slice();
            const original = await FolioDocxReviewer.fromBuffer(base);
            const target = await FolioDocxReviewer.fromBuffer(revised);
            const originalTexts = original.snapshot().blocks.map((block) => block.text);
            // Separate stable-anchor gaps ensure replacement includes a deletion
            // and insertion, rather than silently becoming a paired text edit.
            const events = alignFolioBlocks(original.snapshot().blocks, target.snapshot().blocks);
            expect(events.filter((event) => event.type === "baseOnly")).toHaveLength(
              removal ? 1 : 0,
            );
            expect(events.filter((event) => event.type === "revisedOnly")).toHaveLength(
              independent ? 2 : 1,
            );
            const outcome = await generateRedlineDocx(base, revised).then(
              (value) => ({ status: "returned" as const, value }),
              (error: unknown) => ({ status: "refused" as const, error }),
            );
            if (styleCase === "missingLinkedStyle") {
              expect(applySpy).not.toHaveBeenCalled();
              expect(outcome.status).toBe("refused");
              if (outcome.status !== "refused") panic("A partial redline was returned");
              expect(outcome.error).toBeInstanceOf(GenerateRedlineDocxResourceImportError);
              expect(outcome.error).toMatchObject({
                _tag: "GenerateRedlineDocxResourceImportError",
                detail: "referenced target style MissingStyle is missing",
                message: expect.any(String),
              });
            } else {
              expect(outcome.status).toBe("returned");
              if (outcome.status !== "returned") panic("An unknown direct style was refused");
              expect(outcome.value.skipped).toEqual([]);
              expect(applySpy).toHaveBeenCalledTimes(1);
              const accepting = await FolioDocxReviewer.fromBuffer(outcome.value.buffer);
              expect(accepting.acceptAll()).toBeGreaterThan(0);
              const accepted = await FolioDocxReviewer.fromBuffer(await accepting.toBuffer());
              expect(accepted.snapshot().blocks.map(({ text }) => text)).toEqual(
                target.snapshot().blocks.map(({ text }) => text),
              );
              const replacement = expectParagraphBlock(
                accepted
                  .snapshot()
                  .blocks.find(({ text }) => text === `Unsupported replacement ${suffix}`),
              );
              expect(replacement.styleId).toBe(UNSUPPORTED_STYLE);
              expect(
                accepted
                  .toDocument()
                  .package.styles?.styles.some(({ styleId }) => styleId === UNSUPPORTED_STYLE),
              ).toBe(false);
              const rejecting = await FolioDocxReviewer.fromBuffer(outcome.value.buffer);
              expect(rejecting.rejectAll()).toBeGreaterThan(0);
              const rejected = await FolioDocxReviewer.fromBuffer(await rejecting.toBuffer());
              expect(rejected.snapshot().blocks.map(({ text }) => text)).toEqual(originalTexts);
            }
            expect(new Uint8Array(base)).toEqual(baseBytes);
            expect(new Uint8Array(revised)).toEqual(revisedBytes);
            const reopened = await FolioDocxReviewer.fromBuffer(base);
            expect(reopened.snapshot().blocks.map((block) => block.text)).toEqual(originalTexts);
            expect(reopened.acceptAll()).toBe(0);
            expect(reopened.rejectAll()).toBe(0);
          }
        }
      }),
      // Every generated text exercises all four removal/insertion combinations.
      // Four texts keep the real DOCX parse/serialize surface inexpensive locally.
      { numRuns: 4 },
    );
  } finally {
    applySpy.mockRestore();
  }
});
