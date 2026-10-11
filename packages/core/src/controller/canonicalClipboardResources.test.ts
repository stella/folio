import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { paragraphNumberingReference } from "@stll/docx-core/model";
import { DOCUMENT_OP_REFUSAL_REASONS } from "@stll/docx-core/ops";
import type { Document, Paragraph } from "../types/document";
import { assertExactModel } from "../../../../test/exactModel";
import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import {
  flattenClipboardStyleReferences,
  importClipboardStyles,
} from "./canonicalClipboardResources";

setDefaultTimeout(propertyTestTimeout(30_000));

const fixture = () => {
  const destination = {
    package: {
      document: { content: [] },
      styles: {
        docDefaults: { rPr: { italic: true } },
        styles: [{ styleId: "Foo", type: "paragraph", name: "Owned" }],
      },
    },
  } satisfies Document;
  const source = {
    package: {
      document: { content: [] },
      styles: {
        docDefaults: { rPr: { bold: true } },
        styles: [
          {
            styleId: "Foo",
            type: "paragraph",
            default: true,
            basedOn: "Foo_clipboard1",
            next: "Foo",
            link: "Char",
            pPr: { numPr: paragraphNumberingReference({ numId: 1, ilvl: 0 }) },
          },
          { styleId: "Foo_clipboard1", type: "paragraph", pPr: { keepNext: true } },
          { styleId: "Char", type: "character", link: "Foo" },
          { styleId: "NumStyle", type: "numbering" },
          { styleId: "LevelStyle", type: "paragraph" },
        ],
      },
      numbering: {
        nums: [{ numId: 1, abstractNumId: 1 }],
        abstractNums: [
          {
            abstractNumId: 1,
            numStyleLink: "NumStyle",
            styleLink: "Char",
            levels: [{ ilvl: 0, numFmt: "decimal", lvlText: "%1.", pStyle: "LevelStyle" }],
          },
        ],
      },
    },
  } satisfies Document;
  const paragraphs = [
    {
      type: "paragraph",
      paraId: "00000001",
      formatting: { styleId: "Foo" },
      propertyChanges: [
        {
          type: "paragraphPropertyChange",
          info: { id: 1, author: "Earlier" },
          previousFormatting: { styleId: "Foo_clipboard1" },
        },
      ],
      content: [
        {
          type: "run",
          formatting: { styleId: "Char" },
          content: [{ type: "text", text: "clipboard" }],
        },
      ],
    },
  ] satisfies Paragraph[];
  return { destination, source, paragraphs };
};

test("generated source order preserves deterministic style identities and transitive numbering aliases", () => {
  assertProperty(
    fc.property(fc.shuffledSubarray([0, 1, 2, 3, 4], { minLength: 5, maxLength: 5 }), (order) => {
      const { destination, source, paragraphs } = fixture();
      const snapshot = structuredClone({ destination, source, paragraphs });
      const baseline = importClipboardStyles({ destination, source, paragraphs }).unwrap();
      const reordered = structuredClone(source);
      reordered.package.styles.styles = order.map((index) => {
        const style = source.package.styles.styles.at(index);
        if (!style) throw new TypeError("Generated style order lost a definition.");
        return structuredClone(style);
      });
      const imported = importClipboardStyles({
        destination,
        source: reordered,
        paragraphs,
      }).unwrap();
      assertExactModel(imported, baseline);
      assertExactModel({ destination, source, paragraphs }, snapshot);
      expect([...imported.styleIds.keys()]).toEqual([
        "Char",
        "Foo",
        "Foo_clipboard1",
        "LevelStyle",
        "NumStyle",
      ]);
      expect(imported.styleIds.get("Foo")).toBe("Foo_clipboard2");
      expect([...imported.numberingIds]).toEqual([1]);
      const foo = imported.styles?.styles.find((style) => style.styleId === "Foo_clipboard2");
      expect(foo?.default).toBeUndefined();
      expect(foo?.basedOn).toBe("Foo_clipboard1");
      expect(foo?.next).toBe("Foo_clipboard2");
      const linked = imported.styles?.styles.find((style) => style.styleId === foo?.link);
      expect(linked?.link).toBe("Foo_clipboard2");
      const paragraph = imported.paragraphs.at(0);
      expect(paragraph?.formatting?.styleId).toBe("Foo_clipboard2");
      expect(paragraph?.formatting?.keepNext).toBe(true);
      expect(paragraph?.propertyChanges?.at(0)?.previousFormatting.styleId).toBe("Foo_clipboard1");
      expect(paragraph).not.toBe(paragraphs.at(0));
      assertExactModel(imported.styles?.docDefaults, destination.package.styles.docDefaults);
      assertExactModel(imported.styles?.styles.at(0), destination.package.styles.styles.at(0));
    }),
    {
      numRuns: 25,
      id: "generated source order preserves deterministic style identities and transitive numbering aliases",
    },
  );
});

test("selected basedOn cycles refuse without rewriting the imported source graph", () => {
  const base = fixture();
  const source = {
    ...base.source,
    package: {
      ...base.source.package,
      styles: {
        ...base.source.package.styles,
        styles: base.source.package.styles.styles.map((style) => {
          if (style.styleId === "Foo") return { ...style, basedOn: "Char" };
          if (style.styleId === "Char") return { ...style, basedOn: "Foo" };
          return style;
        }),
      },
    },
  };
  const snapshot = structuredClone({
    source,
    destination: base.destination,
    paragraphs: base.paragraphs,
  });
  const imported = importClipboardStyles({
    source,
    destination: base.destination,
    paragraphs: base.paragraphs,
  });
  expect(imported.isErr()).toBe(true);
  if (imported.isOk()) throw new TypeError("Cyclic source style graph unexpectedly imported.");
  expect(imported.error.reason).toBe(DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH);
  assertExactModel(
    { source, destination: base.destination, paragraphs: base.paragraphs },
    snapshot,
  );
});

test("imported styles and inline formatting resolve against the source theme before destination insertion", () => {
  const base = fixture();
  const source = {
    ...base.source,
    package: {
      ...base.source.package,
      theme: {
        colorScheme: { accent1: "000000", accent2: "FFFFFF" },
        fontScheme: { minorFont: { latin: "Source Serif" } },
      },
      styles: {
        ...base.source.package.styles,
        styles: base.source.package.styles.styles.map((style) => ({
          ...style,
          rPr: {
            fontFamily: { asciiTheme: "minorHAnsi" },
            color: { themeColor: "accent1", themeTint: "80" },
          },
        })),
      },
    },
  } satisfies Document;
  const destination = {
    ...base.destination,
    package: {
      ...base.destination.package,
      theme: {
        colorScheme: { accent1: "FF0000", accent2: "0000FF" },
        fontScheme: { minorFont: { latin: "Destination Sans" } },
      },
    },
  } satisfies Document;
  const paragraphs = base.paragraphs.map(
    (paragraph) =>
      Object.assign({}, paragraph, {
        content: [
          {
            type: "run",
            formatting: {
              fontFamily: { asciiTheme: "minorHAnsi" },
              color: { themeColor: "accent2", themeShade: "80" },
            },
            content: [{ type: "text", text: "themed" }],
          },
        ] satisfies Paragraph["content"],
      }) satisfies Paragraph,
  );
  const snapshot = structuredClone({ source, destination, paragraphs });
  const imported = importClipboardStyles({ source, destination, paragraphs }).unwrap();
  const style = imported.styles?.styles.find((entry) => entry.styleId === "Foo_clipboard2");
  expect(style?.rPr?.fontFamily?.ascii).toBe("Source Serif");
  expect(style?.rPr?.fontFamily?.asciiTheme).toBeUndefined();
  expect(style?.rPr?.color).toEqual({ rgb: "7F7F7F" });
  const run = imported.paragraphs.at(0)?.content.at(0);
  if (run?.type !== "run") throw new TypeError("Themed clipboard run disappeared.");
  expect(run.formatting?.fontFamily?.ascii).toBe("Source Serif");
  expect(run.formatting?.fontFamily?.asciiTheme).toBeUndefined();
  expect(run.formatting?.color).toEqual({ rgb: "808080" });
  assertExactModel({ source, destination, paragraphs }, snapshot);
});

test("source defaults and style toggles materialize beneath direct formatting in nested and empty paragraphs", () => {
  const destination = {
    package: {
      document: { content: [] },
      styles: {
        docDefaults: {
          rPr: { fontFamily: { ascii: "Destination Sans" }, color: { rgb: "FF0000" } },
          pPr: { spaceBefore: 10, spacingExplicit: { before: true, after: true } },
        },
        styles: [],
      },
    },
  } satisfies Document;
  const source = {
    package: {
      document: { content: [] },
      styles: {
        docDefaults: {
          rPr: { fontFamily: { ascii: "Source Serif" }, color: { rgb: "112233" }, bold: false },
          pPr: { spaceBefore: 240 },
        },
        styles: [
          {
            styleId: "Paragraph",
            type: "paragraph",
            rPr: { bold: true },
            pPr: { spaceBefore: 480 },
          },
          { styleId: "Character", type: "character", rPr: { bold: true } },
        ],
      },
    },
  } satisfies Document;
  for (const kind of ["unstyled", "styled", "direct", "empty"] as const) {
    const directParagraphProps = kind === "direct" ? { spaceBefore: 720 } : {};
    const directRunProps = kind === "direct" ? { bold: true } : {};
    const expectedSpacing = { unstyled: 240, styled: 480, direct: 720, empty: 240 };
    const paragraphs = [
      {
        type: "paragraph",
        paraId: "00000001",
        ...(kind === "styled" || kind === "direct"
          ? { formatting: { styleId: "Paragraph", ...directParagraphProps } }
          : {}),
        content:
          kind === "empty"
            ? []
            : [
                {
                  type: "hyperlink",
                  anchor: "bookmark",
                  children: [
                    {
                      type: "run",
                      formatting:
                        kind === "unstyled" ? {} : { styleId: "Character", ...directRunProps },
                      content: [{ type: "text", text: "nested" }],
                    },
                  ],
                },
              ],
      },
    ] satisfies Paragraph[];
    const snapshot = structuredClone({ source, destination, paragraphs });
    const imported = importClipboardStyles({ source, destination, paragraphs }).unwrap();
    const paragraph = imported.paragraphs.at(0);
    expect(paragraph?.formatting?.spaceBefore).toBe(expectedSpacing[kind]);
    expect(paragraph?.formatting?.spacingExplicit).toBeUndefined();
    if (kind === "empty") {
      expect(paragraph?.formatting?.runProperties?.fontFamily?.ascii).toBe("Source Serif");
      expect(paragraph?.formatting?.runProperties?.color?.rgb).toBe("112233");
    } else {
      const hyperlink = paragraph?.content.at(0);
      if (hyperlink?.type !== "hyperlink")
        throw new TypeError("Nested default fixture disappeared.");
      const run = hyperlink.children.at(0);
      if (run?.type !== "run") throw new TypeError("Nested default run disappeared.");
      expect(run.formatting?.fontFamily?.ascii).toBe("Source Serif");
      expect(run.formatting?.color?.rgb).toBe("112233");
      expect(run.formatting?.bold).toBe(kind === "direct");
    }
    assertExactModel(imported.styles?.docDefaults, destination.package.styles.docDefaults);
    assertExactModel({ source, destination, paragraphs }, snapshot);
  }
});

test.each(["font", "colour"] as const)("unresolved source theme %s refuses atomically", (kind) => {
  const base = fixture();
  const paragraphs = base.paragraphs.map(
    (paragraph) =>
      ({
        ...paragraph,
        content: [
          {
            type: "run",
            formatting:
              kind === "font"
                ? { fontFamily: { asciiTheme: "minorHAnsi" } }
                : { color: { themeColor: "accent1" } },
            content: [{ type: "text", text: "unresolved" }],
          },
        ],
      }) satisfies Paragraph,
  );
  const snapshot = structuredClone({ ...base, paragraphs });
  const imported = importClipboardStyles({
    source: base.source,
    destination: base.destination,
    paragraphs,
  });
  expect(imported.isErr()).toBe(true);
  if (imported.isOk()) throw new TypeError("Missing source theme unexpectedly imported.");
  expect(imported.error.reason).toBe(DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH);
  assertExactModel({ ...base, paragraphs }, snapshot);
});

test.each([
  "missingStyle",
  "duplicateStyle",
  "missingNum",
  "missingAbstract",
  "duplicateNum",
  "duplicateAbstract",
] as const)("%s refuses without mutating either package or clipboard reviews", (failure) => {
  const { destination, source, paragraphs } = fixture();
  switch (failure) {
    case "missingStyle":
      source.package.styles.styles = source.package.styles.styles.filter(
        (style) => style.styleId !== "Char",
      );
      break;
    case "duplicateStyle": {
      const style = source.package.styles.styles.at(0);
      if (!style) throw new TypeError("Missing style fixture.");
      source.package.styles.styles.push(structuredClone(style));
      break;
    }
    case "missingNum":
      source.package.numbering.nums = [];
      break;
    case "missingAbstract":
      source.package.numbering.abstractNums = [];
      break;
    case "duplicateNum":
      source.package.numbering.nums.push({ numId: 1, abstractNumId: 1 });
      break;
    case "duplicateAbstract": {
      const abstract = source.package.numbering.abstractNums.at(0);
      if (!abstract) throw new TypeError("Missing abstract fixture.");
      source.package.numbering.abstractNums.push(structuredClone(abstract));
      break;
    }
    default:
      failure satisfies never;
  }
  const snapshot = structuredClone({ destination, source, paragraphs });
  const imported = importClipboardStyles({ destination, source, paragraphs });
  expect(imported.isErr()).toBe(true);
  if (imported.isOk()) throw new TypeError("Invalid source package was imported.");
  expect(imported.error.reason).toBe(DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH);
  assertExactModel({ destination, source, paragraphs }, snapshot);
});

test("normalized foreign style hints flatten every mark and review reference while preserving authored direct formatting", () => {
  const normalized = (styleHint: string | undefined) => {
    const formatting = { alignment: "right", runProperties: { italic: true } } as const;
    const previousParagraph = { alignment: "center" } as const;
    const currentParagraph = { alignment: "right" } as const;
    const runFormatting = { bold: true };
    const previousRun = { italic: false };
    const currentRun = { bold: true };
    if (styleHint !== undefined) {
      for (const record of [
        formatting,
        formatting.runProperties,
        previousParagraph,
        currentParagraph,
        runFormatting,
        previousRun,
        currentRun,
      ])
        Reflect.set(record, "styleId", styleHint);
    }
    return [
      {
        type: "paragraph",
        paraId: "00000001",
        formatting,
        propertyChanges: [
          {
            type: "paragraphPropertyChange",
            info: { id: 10, author: "Source" },
            previousFormatting: previousParagraph,
            currentFormatting: currentParagraph,
          },
        ],
        content: [
          {
            type: "hyperlink",
            href: "https://example.com/source",
            children: [
              {
                type: "run",
                formatting: runFormatting,
                propertyChanges: [
                  {
                    type: "runPropertyChange",
                    info: { id: 11, author: "Source" },
                    previousFormatting: previousRun,
                    currentFormatting: currentRun,
                  },
                ],
                content: [{ type: "text", text: "normalized CSS" }],
              },
            ],
          },
        ],
      },
    ] satisfies Paragraph[];
  };
  const destination = fixture().destination;
  const destinationBefore = structuredClone(destination);
  for (const styleHint of ["Foo", "UnknownForeign"]) {
    const paragraphs = normalized(styleHint);
    flattenClipboardStyleReferences(paragraphs);
    assertExactModel(paragraphs, normalized(undefined));
    assertExactModel(destination, destinationBefore);
  }
});

// Modifiers span all byte values; tint and shade retain opposite channel contributions.
test("every theme modifier byte imports independent grayscale tint and shade values", () => {
  assertProperty(
    fc.property(fc.integer({ min: 0, max: 255 }), (gray) => {
      const channels = [gray, gray, gray];
      const hex = (values: readonly number[]) =>
        values
          .map((value) => value.toString(16).padStart(2, "0"))
          .join("")
          .toUpperCase();
      const rgb = hex(channels);
      for (let byte = 0; byte <= 255; byte += 1) {
        const modifier = byte.toString(16).padStart(2, "0").toUpperCase();
        const source = {
          package: {
            document: { content: [] },
            theme: { colorScheme: { accent1: rgb } },
            styles: {
              styles: [
                {
                  styleId: "Foreign",
                  type: "paragraph",
                  rPr: { color: { themeColor: "accent1", themeTint: modifier } },
                },
              ],
            },
          },
        } satisfies Document;
        const destination = {
          package: {
            document: { content: [] },
            theme: { colorScheme: { accent1: "123456" } },
            styles: { styles: [] },
          },
        } satisfies Document;
        const paragraphs = [
          {
            type: "paragraph",
            paraId: "00000001",
            formatting: { styleId: "Foreign" },
            content: [
              {
                type: "run",
                formatting: { color: { themeColor: "accent1", themeShade: modifier } },
                content: [{ type: "text", text: "color" }],
              },
            ],
          },
        ] satisfies Paragraph[];
        const before = structuredClone({ source, destination, paragraphs });
        const imported = importClipboardStyles({ source, destination, paragraphs }).unwrap();
        const style = imported.styles?.styles.find((entry) => entry.styleId === "Foreign");
        const run = imported.paragraphs.at(0)?.content.at(0);
        if (run?.type !== "run") throw new TypeError("Imported modifier run disappeared.");
        // Grayscale lightness agrees in RGB and HSL; compute without the resolver.
        const tinted = hex(
          channels.map((channel) => Math.round((channel * byte + 255 * (255 - byte)) / 255)),
        );
        const shaded = hex(channels.map((channel) => Math.round((channel * byte) / 255)));
        expect(style?.rPr?.color).toEqual({ rgb: tinted });
        expect(run.formatting?.color).toEqual({ rgb: shaded });
        assertExactModel({ source, destination, paragraphs }, before);
      }
    }),
    {
      numRuns: 10,
      id: "every theme modifier byte imports independent grayscale tint and shade values",
    },
  );
});
