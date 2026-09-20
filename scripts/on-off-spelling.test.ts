/**
 * One spelling of an `ST_OnOff` value per polarity, across every serializer.
 *
 * `ST_OnOff` accepts `1`/`true`/`on` and `0`/`false`/`off`, and folio's readers
 * take all six: a source that spelled an off as `off` reads as an off, and a
 * captured element replays the bytes it arrived as. What folio *writes* is one
 * spelling, `0` for an off and the bare element for an on, because a package
 * that writes three spellings of the same value makes every byte-level
 * comparison — a save fixed point, a corpus diff, a reviewer's eye — argue
 * about which of them it is looking at.
 *
 * The scan is over the emitted markup rather than over the readers, so a
 * parser's documentation of the lexical space is not a violation of it.
 */

import { describe, expect, test } from "bun:test";
import path from "node:path";

const PACKAGES_DIR = path.resolve(import.meta.dir, "..", "packages");

/**
 * A literal element with a `w:val` in a spelling folio does not write.
 *
 * Anchored on the element so that a reader's prose about `w:val="off"` does
 * not match; only markup a serializer hands to a consumer does.
 */
const NON_CANONICAL_ON_OFF = /<w:[A-Za-z]+[^>]*\sw:val="(?:on|off|true|false)"/u;

/**
 * Nothing is exempt. An entry here would be a second spelling in the tree,
 * which is the thing this test exists to refuse.
 */
const ALLOWED: ReadonlySet<string> = new Set();

const isScanned = (file: string): boolean =>
  !file.includes("/__tests__/") &&
  !file.endsWith(".test.ts") &&
  !file.endsWith(".test.tsx") &&
  !file.includes("/__fixtures__/");

describe("ST_OnOff spelling", () => {
  test("no serializer writes an on/off value as a word", async () => {
    const offenders: string[] = [];
    for await (const file of new Bun.Glob("*/src/**/*.ts").scan({
      absolute: true,
      cwd: PACKAGES_DIR,
    })) {
      const relative = path.relative(PACKAGES_DIR, file);
      if (!isScanned(`/${relative}`) || ALLOWED.has(relative)) {
        continue;
      }
      const source = await Bun.file(file).text();
      for (const [index, line] of source.split("\n").entries()) {
        if (NON_CANONICAL_ON_OFF.test(line)) {
          offenders.push(`${relative}:${index + 1}`);
        }
      }
    }
    expect(offenders).toEqual([]);
  });

  test("the scan sees the spellings it refuses", () => {
    // Anti-vacuity: a pattern that matched nothing would pass the property
    // above for the wrong reason, and so would one that stopped reaching the
    // tree. Both polarities, both a bare element and one with neighbours.
    expect(NON_CANONICAL_ON_OFF.test('<w:hideMark w:val="off"/>')).toBe(true);
    expect(NON_CANONICAL_ON_OFF.test('<w:updateFields w:val="true"/>')).toBe(true);
    expect(NON_CANONICAL_ON_OFF.test('<w:top w:sz="4" w:val="on"/>')).toBe(true);
    expect(NON_CANONICAL_ON_OFF.test('<w:hidden w:val="0"/>')).toBe(false);
    expect(NON_CANONICAL_ON_OFF.test("<w:hidden/>")).toBe(false);
    // A reader's prose about the lexical space is not markup.
    expect(NON_CANONICAL_ON_OFF.test(' * - w:val="false", w:val="0", or w:val="off"')).toBe(false);
  });
});
