import { expect, setDefaultTimeout, spyOn, test } from "bun:test";
import { panic } from "better-result";
import fc from "fast-check";

import { assertProperty, propertyTestTimeout } from "../../../test/property-testing";
import { expectParagraphBlock } from "../../../test/paragraphBlock";
import { FolioDocxReviewer } from "./ai-edits/headless";
import { createDocx } from "./docx/rezip";
import { generateRedlineDocx, GenerateRedlineDocxResourceImportError } from "./redline";
import { createStyleResolver } from "./prosemirror/styles/styleResolver";
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
      {
        id: "a failing insertion style closure refuses the whole redline without changing input packages",
        numRuns: 4,
      },
    );
  } finally {
    applySpy.mockRestore();
  }
});

test("unstyled redline insertions preserve differing default paragraph styles in equal contexts", async () => {
  await assertProperty(
    fc.asyncProperty(
      fc.record({
        fontSize: fc.integer({ min: 16, max: 40 }),
        spaceAfter: fc.integer({ min: 0, max: 400 }),
        suffix: fc.stringMatching(/^[a-z]{1,12}$/u),
      }),
      async ({ fontSize, spaceAfter, suffix }) => {
        for (const sameStyleId of [false, true]) {
          for (const sourceBold of [false, true]) {
            const fixture = async (side: "base" | "revised") => {
              const document = createEmptyDocument();
              document.package.theme = {
                fontScheme: { minorFont: { latin: "Shared Font" } },
              };
              document.package.styles = {
                docDefaults: { rPr: { fontSize } },
                styles: [
                  {
                    type: "paragraph",
                    styleId: side === "base" || sameStyleId ? NORMAL_STYLE : "TargetDefault",
                    default: true,
                    pPr: { spaceAfter: side === "base" ? spaceAfter : spaceAfter + 20 },
                    rPr: { bold: side === "revised" ? sourceBold : !sourceBold },
                  },
                  { type: "paragraph", styleId: "Anchor" },
                ],
              };
              document.package.document.content = [
                paragraph("93000001", `Leading anchor ${suffix}`, "Anchor"),
                ...(side === "revised"
                  ? [{ ...paragraph("93000002", `Unstyled insertion ${suffix}`), formatting: {} }]
                  : []),
                paragraph("93000003", `Trailing anchor ${suffix}`, "Anchor"),
              ];
              return createDocx(document);
            };
            const base = await fixture("base");
            const revised = await fixture("revised");
            const original = await FolioDocxReviewer.fromBuffer(base);
            const target = await FolioDocxReviewer.fromBuffer(revised);
            const effectiveFormatting = (reviewer: FolioDocxReviewer) => {
              const resolver = createStyleResolver(reviewer.toDocument().package.styles);
              return reviewer.snapshot().blocks.map((block) => {
                const paragraphBlock = expectParagraphBlock(block);
                return {
                  text: paragraphBlock.text,
                  formatting: resolver.resolveParagraphStyle(paragraphBlock.styleId),
                };
              });
            };
            const outcome = await generateRedlineDocx(base, revised);
            expect(outcome.skipped).toEqual([]);
            const accepting = await FolioDocxReviewer.fromBuffer(outcome.buffer);
            expect(accepting.acceptAll()).toBeGreaterThan(0);
            const accepted = await FolioDocxReviewer.fromBuffer(await accepting.toBuffer());
            expect(effectiveFormatting(accepted)).toEqual(effectiveFormatting(target));
            const inserted = expectParagraphBlock(accepted.snapshot().blocks.at(1));
            expect(inserted.styleId).toBeDefined();
            const rejecting = await FolioDocxReviewer.fromBuffer(outcome.buffer);
            expect(rejecting.rejectAll()).toBeGreaterThan(0);
            const rejected = await FolioDocxReviewer.fromBuffer(await rejecting.toBuffer());
            expect(effectiveFormatting(rejected)).toEqual(effectiveFormatting(original));
          }
        }
      },
    ),
    {
      id: "unstyled redline insertions preserve differing default paragraph styles in equal contexts",
      numRuns: 4,
    },
  );
});
