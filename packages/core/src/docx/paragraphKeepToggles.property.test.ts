/**
 * A paragraph toggle a command sets reaches `w:pPr`.
 *
 * `PARAGRAPH_FORMATTING_WRITE_BACK` is total over `ParagraphFormatting`, so
 * every `w:pPr` field must state how a save tells its authored value from its
 * inherited one. `keepNext`, `keepLines` and `runInWithNext` were classified
 * `original-only`: an imported value survived through
 * `ParagraphAttrs._originalFormatting`, and a value a command set had nowhere
 * to go. A paragraph with no `w:pPr` of its own silently lost every one of
 * them on save.
 *
 * They are `style-resolved-attr` now, like `widowControl`, which is the
 * disposition their attrs already deserved: `toProseDoc` seeds them with the
 * value the style cascade resolves to. The property therefore has to pin both
 * halves at once, over every combination of the three toggles crossed with
 * what the paragraph's style says:
 *
 *   authored   a value the command set is written to `w:pPr` and read back.
 *   inherited  a value that only echoes the style is NOT written, so a later
 *              edit to that style still reaches the paragraph.
 *
 * `false` is a decision, not an absence: `w:keepNext w:val="0"` cancels a
 * style's `keepNext` and must survive as itself.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { Document, ParagraphFormatting, Style } from "../types/document";
import { parseDocx } from "./parser";
import { createEmptyDocx, repackDocx } from "./rezip";

const TOGGLES = ["keepNext", "keepLines", "runInWithNext"] as const;
type Toggle = (typeof TOGGLES)[number];

/** The tri-state of an OOXML toggle, plus what the style says about it. */
const TOGGLE_VALUES = [undefined, true, false] as const;

type Toggles = Record<Toggle, boolean | undefined>;

type ToggleCase = {
  /** What the source paragraph states in its own `w:pPr`. */
  paragraph: Toggles;
  /** What its style says. */
  style: Toggles;
  /** What a command then sets on the paragraph node. */
  commanded: Toggles;
};

/**
 * What `w:pPr` must carry after the command.
 *
 * An undecided attr clears the field. A paragraph that already stated the
 * field keeps stating it. Otherwise a value that only echoes the style is
 * withheld, so a later edit to that style still reaches the paragraph.
 */
const expectedDirect = (
  commanded: boolean | undefined,
  style: boolean | undefined,
  authored: boolean | undefined,
): boolean | undefined => {
  if (commanded === undefined) {
    return undefined;
  }
  if (authored !== undefined) {
    return commanded;
  }
  return commanded === style ? undefined : commanded;
};

const STYLE_ID = "KeepStyle";

/**
 * `runInWithNext` is `<w:specVanish/>` on the paragraph mark. Like the other
 * toggles, its direct value is tri-state: absent, on, or an explicit off that
 * cancels an inherited value.
 */
const toggleRecord = fc.record({
  keepNext: fc.constantFrom(...TOGGLE_VALUES),
  keepLines: fc.constantFrom(...TOGGLE_VALUES),
  runInWithNext: fc.constantFrom(...TOGGLE_VALUES),
});

/**
 * A style states `w:specVanish` in its own `w:rPr`, which folio's style model
 * does not project onto `paragraphFormatting`, so `runInWithNext` has no style
 * value to inherit and the style leg leaves it undecided.
 */
const styleToggleRecord = fc.record({
  keepNext: fc.constantFrom(...TOGGLE_VALUES),
  keepLines: fc.constantFrom(...TOGGLE_VALUES),
  runInWithNext: fc.constant(undefined),
});

const toggleCase = fc.record({
  paragraph: toggleRecord,
  style: styleToggleRecord,
  commanded: toggleRecord,
});

const formattingFrom = (toggles: Toggles): ParagraphFormatting => {
  const formatting: ParagraphFormatting = {};
  for (const toggle of TOGGLES) {
    const value = toggles[toggle];
    if (value !== undefined) {
      formatting[toggle] = value;
    }
  }
  return formatting;
};

const styleWith = (toggles: Toggles): Style => ({
  type: "paragraph",
  styleId: STYLE_ID,
  name: "Keep Style",
  pPr: formattingFrom(toggles),
});

const withParagraph = (document: Document, testCase: ToggleCase): Document => {
  const styles = document.package.styles ?? { styles: [] };
  return {
    ...document,
    package: {
      ...document.package,
      styles: {
        ...styles,
        styles: [
          ...styles.styles.filter((style) => style.styleId !== STYLE_ID),
          styleWith(testCase.style),
        ],
      },
      document: {
        ...document.package.document,
        content: [
          {
            type: "paragraph",
            formatting: { styleId: STYLE_ID, ...formattingFrom(testCase.paragraph) },
            content: [{ type: "run", content: [{ type: "text", text: "kept" }] }],
          },
        ],
      },
    },
  };
};

/** What the paragraph states itself, which is all `w:pPr` may carry. */
const readDirect = (document: Document): Toggles => {
  const block = document.package.document.content.at(0);
  const formatting = block?.type === "paragraph" ? block.formatting : undefined;
  return {
    keepNext: formatting?.keepNext,
    keepLines: formatting?.keepLines,
    runInWithNext: formatting?.runInWithNext,
  };
};

const parse = (buffer: ArrayBuffer): Promise<Document> =>
  parseDocx(buffer, { detectVariables: false, preloadFonts: false });

const save = (document: Document): Promise<ArrayBuffer> =>
  repackDocx(document, { updateModifiedDate: false });

/** Set the toggles the way an editor command would: on the paragraph node. */
const setTogglesThroughTheEditor = (document: Document, toggles: Toggles): Document => {
  const proseDoc = toProseDoc(document);
  const paragraph = proseDoc.child(0);
  const next = proseDoc.copy(
    proseDoc.content.replaceChild(
      0,
      paragraph.type.create({ ...paragraph.attrs, ...toggles }, paragraph.content, paragraph.marks),
    ),
  );
  return fromProseDoc(next, document);
};

describe("keep toggles a command sets reach w:pPr", () => {
  test("a keepNext set on a paragraph with no w:pPr of its own survives the save", async () => {
    const template = await parse(await createEmptyDocx());
    const plain = await parse(
      await save({
        ...template,
        package: {
          ...template.package,
          document: {
            ...template.package.document,
            content: [
              {
                type: "paragraph",
                content: [{ type: "run", content: [{ type: "text", text: "p" }] }],
              },
            ],
          },
        },
      }),
    );
    expect(readDirect(plain)).toEqual({
      keepNext: undefined,
      keepLines: undefined,
      runInWithNext: undefined,
    });

    const edited = setTogglesThroughTheEditor(plain, {
      keepNext: true,
      keepLines: true,
      runInWithNext: undefined,
    });
    expect(readDirect(await parse(await save(edited)))).toEqual({
      keepNext: true,
      keepLines: true,
      runInWithNext: undefined,
    });
  });

  test(
    "each toggle is written when the paragraph states it and withheld when the style does",
    async () => {
      const template = await parse(await createEmptyDocx());

      await fc.assert(
        fc.asyncProperty(toggleCase, async (testCase) => {
          const opened = await parse(await save(withParagraph(template, testCase)));
          expect(readDirect(opened)).toEqual(testCase.paragraph);

          // Untouched: the editor round trip keeps the paragraph's own values
          // and does not materialise the style's.
          const roundTripped = await parse(await save(fromProseDoc(toProseDoc(opened), opened)));
          expect(readDirect(roundTripped)).toEqual(testCase.paragraph);

          // Commanded: every toggle the command states lands in `w:pPr`,
          // except one that only repeats what the style already says.
          const commanded = await parse(
            await save(setTogglesThroughTheEditor(opened, testCase.commanded)),
          );
          const direct = readDirect(commanded);
          for (const toggle of TOGGLES) {
            expect(direct[toggle]).toBe(
              expectedDirect(
                testCase.commanded[toggle],
                testCase.style[toggle],
                testCase.paragraph[toggle],
              ),
            );
          }
        }),
        propertyConfig({ numRuns: 50 }),
      );
    },
    propertyTestTimeout(90_000),
  );
});
