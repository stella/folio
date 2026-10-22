import { panic } from "better-result";
import fc from "fast-check";

export const TEXT_BOUNDARIES = [7, 4096, 16 * 1024, 64 * 1024] as const;
export const astralText = fc
  .integer({ min: 0x10000, max: 0x10ffff })
  .map((point) => String.fromCodePoint(point));
export const combiningText: fc.Arbitrary<string> = fc.constantFrom(
  "e\u0301",
  "a\u0308\u0323",
  "क्ष",
);
export const rtlText: fc.Arbitrary<string> = fc.constantFrom(
  "שָׁלוֹם",
  "مرحبا",
);
export const edgeText = fc.oneof(
  astralText,
  combiningText,
  rtlText,
  fc.constantFrom("👩🏽‍⚖️", "🇨🇿"),
);

export const boundaryOffset = fc
  .tuple(
    fc.constantFrom(...TEXT_BOUNDARIES),
    fc.integer({ min: 1, max: 4 }),
    fc.integer({ min: -2, max: 1 }),
  )
  .map(([size, multiple, delta]) => size * multiple + delta);

/** ASCII padding positions the edge in both UTF-16 units and UTF-8 bytes. */
export const textAtBoundary = fc
  .tuple(boundaryOffset, edgeText, fc.array(edgeText, { maxLength: 3 }))
  .map(
    ([offset, edge, suffix]) =>
      "a".repeat(offset) + edge + suffix.join("") + "tail",
  );

/** Exercise every serializer at each declared boundary, rather than sampling the set. */
export const boundaryTextCases = (astral: string, suffix: string) =>
  TEXT_BOUNDARIES.flatMap((boundary) =>
    [-1, 0, 1].flatMap((shift) =>
      [astral, "e\u0301", "שָׁלוֹם", "👩🏽‍⚖️", "🇨🇿"].map((cluster) => ({
        boundary,
        text:
          "a".repeat(boundary + shift) +
          cluster +
          astral +
          "e\u0301שלום" +
          suffix +
          "tail",
      })),
    ),
  );

type AssertTextRoundTripOptions = {
  text: string;
  roundTrip: (text: string) => string | Promise<string>;
};

/** Compare the actual producer/reader output with the exact input text. */
export const assertTextRoundTrip = async ({
  text,
  roundTrip,
}: AssertTextRoundTripOptions): Promise<void> => {
  if ((await roundTrip(text)) !== text)
    panic("Text round trip changed Unicode content");
};
