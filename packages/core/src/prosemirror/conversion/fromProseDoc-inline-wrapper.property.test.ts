/**
 * What the save leg writes back around a transparent inline wrapper.
 *
 * `toProseDoc` lifts `w:bdo`/`w:dir` out of the content tree and records the
 * nesting on the `inlineWrapper` mark of the leaves the wrapper held.
 * `fromProseDoc` reads that mark back: it cuts the paragraph's inline sequence
 * into maximal groups of equal stack and closes the wrappers around each
 * group, so an edited span keeps the wrapper it was authored inside instead of
 * losing it with the paragraph the editor rebuilt.
 *
 * Before the save leg read the mark, the wrapper only survived where the
 * source paragraph's markup was replayed, and every round trip below came back
 * as a bare run.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig, propertyTestTimeout } from "../../../../../test/property-testing";

import type {
  Document,
  InlineWrapper,
  Paragraph,
  ParagraphContent,
  TrackedChangeInfo,
} from "../../types/document";
import { createEmptyDocument } from "../../utils/createDocument";
import { fromProseDoc } from "./fromProseDoc";
import { toProseDoc } from "./toProseDoc";

const TEXT = "x";
const RUN = { type: "run", content: [{ type: "text", text: TEXT }] } as const;
const INFO: TrackedChangeInfo = { id: 1, author: "Reviewer", date: "2026-01-01T00:00:00Z" };

const documentWith = (content: Paragraph["content"]): Document => {
  const template = createEmptyDocument();
  return {
    ...template,
    package: {
      ...template.package,
      document: {
        ...template.package.document,
        content: [{ type: "paragraph", paraId: "C0000001", content }],
      },
    },
  };
};

const paragraphContentOf = (document: Document): ParagraphContent[] => {
  const block = document.package.document.content.at(0);
  if (block?.type !== "paragraph") {
    throw new Error("The rebuilt document lost its paragraph");
  }
  return block.content;
};

/** One pass through the editor: model to ProseMirror and back. */
const roundTrip = (content: Paragraph["content"]): ParagraphContent[] => {
  const source = documentWith(content);
  return paragraphContentOf(fromProseDoc(toProseDoc(source), source));
};

/**
 * A run comes back carrying the formatting the paragraph resolved onto it, so
 * the comparison is over the shape the wrappers make and the text they hold,
 * not over every run property the save leg is entitled to write.
 */
type WrapperShape =
  | { kind: "text"; text: string }
  | { kind: "wrapper"; control: string; direction: string | undefined; content: WrapperShape[] }
  | { kind: "revision"; type: string; content: WrapperShape[] }
  | { kind: "hyperlink"; content: WrapperShape[] }
  | { kind: "other"; type: string };

const shapeOf = (item: ParagraphContent): WrapperShape => {
  switch (item.type) {
    case "run":
      return {
        kind: "text",
        text: item.content.map((part) => (part.type === "text" ? part.text : "")).join(""),
      };
    case "inlineWrapper":
      return {
        kind: "wrapper",
        control: item.control,
        direction: item.direction,
        content: item.content.map(shapeOf),
      };
    case "insertion":
    case "deletion":
    case "moveFrom":
    case "moveTo":
      return { kind: "revision", type: item.type, content: item.content.map(shapeOf) };
    case "hyperlink":
      return { kind: "hyperlink", content: item.children.map(shapeOf) };
    default:
      return { kind: "other", type: item.type };
  }
};

const shapesOf = (content: readonly ParagraphContent[]): WrapperShape[] => content.map(shapeOf);

/** An authored wrapper, without the content it holds. */
type AuthoredWrapper = Omit<InlineWrapper, "content">;

const wrapperArbitrary: fc.Arbitrary<AuthoredWrapper> = fc
  .record({
    control: fc.constantFrom("override" as const, "embedding" as const),
    direction: fc.constantFrom("ltr" as const, "rtl" as const, undefined),
  })
  .map(({ control, direction }) =>
    direction === undefined
      ? { type: "inlineWrapper", kind: "bidi", control }
      : { type: "inlineWrapper", kind: "bidi", control, direction },
  );

const HYPERLINK = {
  type: "hyperlink",
  url: "https://example.invalid/",
  children: [RUN],
} as const satisfies ParagraphContent;

const INSERTED_RUN = { type: "insertion", info: INFO, content: [RUN] } as const;

/**
 * The content a wrapper holds.
 *
 * A revision-wrapped run is generated in the authored order the canonical save
 * order preserves — revision inside the wrapper is normalised to revision
 * outside it, which {@link describe} "a wrapper authored outside a revision"
 * asserts on its own.
 */
const contentArbitrary: fc.Arbitrary<ParagraphContent> = fc.constantFrom<ParagraphContent>(
  RUN,
  HYPERLINK,
  INSERTED_RUN,
);

/** `inner` wrapped by `authored`, outermost first. */
const nest = (authored: readonly AuthoredWrapper[], inner: ParagraphContent): ParagraphContent => {
  let nested = inner;
  for (const wrapper of authored.toReversed()) {
    nested = { ...wrapper, content: [nested] };
  }
  return nested;
};

/** The canonical save order: the revision outside the wrappers it was inside. */
const canonical = (
  authored: readonly AuthoredWrapper[],
  inner: ParagraphContent,
): ParagraphContent =>
  inner.type === "insertion" ? { ...inner, content: [nest(authored, RUN)] } : nest(authored, inner);

describe("a wrapper tree the editor gives back", () => {
  test(
    "comes back as the authored nesting",
    () => {
      fc.assert(
        fc.property(
          fc.array(wrapperArbitrary, { minLength: 1, maxLength: 3 }),
          contentArbitrary,
          (authored, inner) => {
            expect(shapesOf(roundTrip([nest(authored, inner)]))).toEqual(
              shapesOf([canonical(authored, inner)]),
            );
          },
        ),
        propertyConfig({ numRuns: 200 }),
      );
    },
    propertyTestTimeout(),
  );

  test(
    "is a fixed point of a second round trip",
    () => {
      fc.assert(
        fc.property(
          fc.array(wrapperArbitrary, { minLength: 1, maxLength: 3 }),
          contentArbitrary,
          (authored, inner) => {
            const once = roundTrip([nest(authored, inner)]);
            expect(JSON.stringify(shapesOf(once))).toContain('"wrapper"');
            expect(shapesOf(roundTrip(once))).toEqual(shapesOf(once));
          },
        ),
        propertyConfig({ numRuns: 200 }),
      );
    },
    propertyTestTimeout(),
  );

  test("keeps two adjacent wrappers apart", () => {
    const left: ParagraphContent = {
      type: "inlineWrapper",
      kind: "bidi",
      control: "override",
      direction: "rtl",
      content: [{ type: "run", content: [{ type: "text", text: "a" }] }],
    };
    const right: ParagraphContent = {
      type: "inlineWrapper",
      kind: "bidi",
      control: "embedding",
      direction: "ltr",
      content: [{ type: "run", content: [{ type: "text", text: "b" }] }],
    };
    expect(shapesOf(roundTrip([left, right]))).toEqual(shapesOf([left, right]));
  });

  test("joins two runs the same wrapper held into one wrapper", () => {
    const wrapped: ParagraphContent = {
      type: "inlineWrapper",
      kind: "bidi",
      control: "override",
      direction: "rtl",
      content: [
        { type: "run", content: [{ type: "text", text: "a" }] },
        { type: "run", content: [{ type: "text", text: "b" }], formatting: { bold: true } },
      ],
    };
    const saved = roundTrip([wrapped]);
    expect(saved).toHaveLength(1);
    expect(saved.at(0)?.type).toBe("inlineWrapper");
  });

  test("leaves text outside the wrapper outside it", () => {
    const outer = { type: "run", content: [{ type: "text", text: "out" }] } as const;
    const wrapped: ParagraphContent = {
      type: "inlineWrapper",
      kind: "bidi",
      control: "embedding",
      direction: "rtl",
      content: [{ type: "run", content: [{ type: "text", text: "in" }] }],
    };
    expect(shapesOf(roundTrip([outer, wrapped, outer]))).toEqual(shapesOf([outer, wrapped, outer]));
  });
});

describe("a wrapper authored outside a revision", () => {
  const WRAPPER_OUTSIDE: ParagraphContent = {
    type: "inlineWrapper",
    kind: "bidi",
    control: "override",
    direction: "rtl",
    content: [INSERTED_RUN],
  };

  const REVISION_OUTSIDE: ParagraphContent = {
    type: "insertion",
    info: INFO,
    content: [
      {
        type: "inlineWrapper",
        kind: "bidi",
        control: "override",
        direction: "rtl",
        content: [RUN],
      },
    ],
  };

  test("is saved with the revision outermost", () => {
    expect(shapesOf(roundTrip([WRAPPER_OUTSIDE]))).toEqual(shapesOf([REVISION_OUTSIDE]));
  });

  test("and one authored inside it save the same way", () => {
    expect(shapesOf(roundTrip([WRAPPER_OUTSIDE]))).toEqual(shapesOf(roundTrip([REVISION_OUTSIDE])));
  });
});

describe("a wrapper with nothing left inside it", () => {
  test("is not written when its only run is gone", () => {
    const empty: ParagraphContent = {
      type: "inlineWrapper",
      kind: "bidi",
      control: "override",
      direction: "rtl",
      content: [],
    };
    expect(roundTrip([empty, RUN])).toEqual(roundTrip([RUN]));
  });
});
