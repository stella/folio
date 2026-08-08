/**
 * Cursive joining decides whether a run boundary severed a letter connection,
 * so the painter and measurer can repair the same boundaries. A wrong verdict
 * either leaves an Arabic word visibly broken or inserts a joiner where none
 * belongs, so the classes and the transparent-skipping rule both need pinning.
 *
 * Characters are written as `\u` escapes throughout: a pasted glyph once
 * silently corrupted a character class in this package's RTL code.
 */

import { describe, expect, test } from "bun:test";

import {
  hasCursiveLetter,
  isCursiveLetter,
  isJoiningTransparent,
  joinsAcrossBoundary,
  joinsToFollowing,
  joinsToPreceding,
} from "./cursiveJoining";
import {
  JOINING_LETTER_RANGES,
  JOINING_TRANSPARENT_RANGES,
  JOINS_BACKWARD_RANGES,
  JOINS_FORWARD_RANGES,
} from "./joiningTypes.gen";

// Arabic letters, by Joining_Type.
const BEH = "ب"; // D — dual-joining
const TEH = "ت"; // D
const KAF = "ك"; // D
const MEEM = "م"; // D
const ALEF = "ا"; // R — right-joining: never connects to what follows
const DAL = "د"; // R
const WAW = "و"; // R
const HAMZA = "ء"; // U — non-joining
// Transparent marks (General_Category Mn, so Joining_Type T by default).
const FATHA = "َ";
const SHADDA = "ّ";
const SUKUN = "ْ";
const ZWJ = "‍"; // C — join-causing
const ZWNJ = "‌"; // U — non-joining, and deliberately so

/** "مكتب" (office): four letters, all connected. */
const CONNECTED_WORD = MEEM + KAF + TEH + BEH;

const allRanges = {
  forward: JOINS_FORWARD_RANGES,
  backward: JOINS_BACKWARD_RANGES,
  transparent: JOINING_TRANSPARENT_RANGES,
  letters: JOINING_LETTER_RANGES,
};

describe("generated joining tables", () => {
  // A malformed table would make every predicate below quietly wrong, so the
  // shape is asserted rather than assumed.
  test.each(Object.entries(allRanges))(
    "%s is well-formed sorted non-overlapping pairs",
    (_name, ranges) => {
      expect(ranges.length % 2).toBe(0);
      for (let i = 0; i < ranges.length; i += 2) {
        expect(ranges[i]!).toBeLessThanOrEqual(ranges[i + 1]!);
        if (i > 0) {
          // Strictly increasing with a gap: the generator coalesces adjacent
          // ranges, so `previousEnd + 1 === nextStart` must not survive.
          expect(ranges[i]!).toBeGreaterThan(ranges[i - 1]! + 1);
        }
      }
    },
  );

  test("a code point is never both a cursive letter and transparent", () => {
    for (let i = 0; i < JOINING_LETTER_RANGES.length; i += 2) {
      for (let cp = JOINING_LETTER_RANGES[i]!; cp <= JOINING_LETTER_RANGES[i + 1]!; cp++) {
        expect(isJoiningTransparent(cp)).toBe(false);
      }
    }
  });
});

describe("joining classes", () => {
  test("dual-joining letters connect in both directions", () => {
    for (const letter of [BEH, TEH, KAF, MEEM]) {
      const cp = letter.codePointAt(0)!;
      expect(joinsToFollowing(cp)).toBe(true);
      expect(joinsToPreceding(cp)).toBe(true);
      expect(isCursiveLetter(cp)).toBe(true);
    }
  });

  test("right-joining letters connect backwards only", () => {
    for (const letter of [ALEF, DAL, WAW]) {
      const cp = letter.codePointAt(0)!;
      expect(joinsToFollowing(cp)).toBe(false);
      expect(joinsToPreceding(cp)).toBe(true);
      expect(isCursiveLetter(cp)).toBe(true);
    }
  });

  test("non-joining characters connect in neither direction", () => {
    for (const ch of [HAMZA, " ", "a", "一", "1", ZWNJ]) {
      const cp = ch.codePointAt(0)!;
      expect(joinsToFollowing(cp)).toBe(false);
      expect(joinsToPreceding(cp)).toBe(false);
    }
  });

  test("ZWJ is join-causing but is not a letter", () => {
    const cp = ZWJ.codePointAt(0)!;
    expect(joinsToFollowing(cp)).toBe(true);
    expect(joinsToPreceding(cp)).toBe(true);
    expect(isCursiveLetter(cp)).toBe(false);
  });

  test("combining marks are transparent, not letters", () => {
    for (const mark of [FATHA, SHADDA, SUKUN]) {
      const cp = mark.codePointAt(0)!;
      expect(isJoiningTransparent(cp)).toBe(true);
      expect(isCursiveLetter(cp)).toBe(false);
    }
  });

  test("hasCursiveLetter gates the Latin hot path", () => {
    expect(hasCursiveLetter("plain ascii")).toBe(false);
    expect(hasCursiveLetter("smlouva o dílo")).toBe(false);
    expect(hasCursiveLetter("日本語")).toBe(false);
    expect(hasCursiveLetter(CONNECTED_WORD)).toBe(true);
    expect(hasCursiveLetter("case " + BEH)).toBe(true);
    // ZWJ alone is join-causing but is not a script, so it must not open the path.
    expect(hasCursiveLetter(ZWJ)).toBe(false);
  });
});

describe("joinsAcrossBoundary", () => {
  test("every internal split of a fully connected word joins", () => {
    for (let cut = 1; cut < CONNECTED_WORD.length; cut++) {
      expect(joinsAcrossBoundary(CONNECTED_WORD.slice(0, cut), CONNECTED_WORD.slice(cut))).toBe(
        true,
      );
    }
  });

  test("a right-joining letter before the cut does not join forwards", () => {
    // "كاتب" (writer): alef sits second and joins only backwards,
    // so the cut after it is the one split in this word that needs no joiner.
    const word = KAF + ALEF + TEH + BEH;
    expect(joinsAcrossBoundary(word.slice(0, 1), word.slice(1))).toBe(true);
    expect(joinsAcrossBoundary(word.slice(0, 2), word.slice(2))).toBe(false);
    expect(joinsAcrossBoundary(word.slice(0, 3), word.slice(3))).toBe(true);
  });

  test("word boundaries and non-joining scripts never join", () => {
    expect(joinsAcrossBoundary(CONNECTED_WORD, " " + CONNECTED_WORD)).toBe(false);
    expect(joinsAcrossBoundary(CONNECTED_WORD + " ", CONNECTED_WORD)).toBe(false);
    expect(joinsAcrossBoundary("bo", "ld")).toBe(false);
    expect(joinsAcrossBoundary("日", "本")).toBe(false);
    expect(joinsAcrossBoundary(CONNECTED_WORD, "1")).toBe(false);
    expect(joinsAcrossBoundary(MEEM, HAMZA)).toBe(false);
  });

  test("empty sides never join", () => {
    expect(joinsAcrossBoundary("", CONNECTED_WORD)).toBe(false);
    expect(joinsAcrossBoundary(CONNECTED_WORD, "")).toBe(false);
    expect(joinsAcrossBoundary("", "")).toBe(false);
  });

  test("ZWNJ is respected: it exists to suppress a join", () => {
    expect(joinsAcrossBoundary(MEEM + ZWNJ, KAF)).toBe(false);
    expect(joinsAcrossBoundary(MEEM, ZWNJ + KAF)).toBe(false);
  });

  // The invariant most likely to regress: vocalized Arabic puts combining marks
  // exactly where the cut lands, and they must be invisible to the verdict.
  test("any number of transparent marks at the boundary cannot change the verdict", () => {
    const marks = [FATHA, SHADDA, SUKUN];
    for (const joined of [true, false]) {
      const before = joined ? MEEM : ALEF;
      const after = KAF;
      for (let count = 0; count <= marks.length; count++) {
        const padding = marks.slice(0, count).join("");
        expect(joinsAcrossBoundary(before + padding, after)).toBe(joined);
        expect(joinsAcrossBoundary(before, padding + after)).toBe(joined);
        expect(joinsAcrossBoundary(before + padding, padding + after)).toBe(joined);
      }
    }
  });

  test("a side made only of transparent marks looks through to nothing", () => {
    // Not "joins to whatever preceded the marks": this helper only sees the two
    // strings it is given, and a caller must pass whole run text.
    expect(joinsAcrossBoundary(FATHA, KAF)).toBe(false);
    expect(joinsAcrossBoundary(MEEM, FATHA)).toBe(false);
  });

  test("astral cursive letters are not split mid-surrogate", () => {
    // Adlam (U+1E900..) is dual-joining and lives above the BMP, so the
    // backwards scan must decode the surrogate pair rather than the low unit.
    const adlamA = String.fromCodePoint(0x1e9_00);
    const adlamB = String.fromCodePoint(0x1e9_01);
    expect(isCursiveLetter(0x1e9_00)).toBe(true);
    expect(joinsAcrossBoundary(adlamA, adlamB)).toBe(true);
    expect(joinsAcrossBoundary(adlamA + FATHA, adlamB)).toBe(true);
  });
});
