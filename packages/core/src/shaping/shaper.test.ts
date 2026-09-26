/**
 * Loading the artifact from bytes, and the bidirectional algorithm it carries.
 */

import { describe, expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import path from "node:path";

import { BIDI_DIRECTION, type BidiDirection, getShaper } from "./shaper";

const SHAPER_MODULE = path.join(import.meta.dir, "shaper.ts");
const WASM_PATH = path.join(import.meta.dir, "..", "generated", "text_shaper_bg.wasm");
const CONFORMANCE_FIXTURE = path.join(
  import.meta.dir,
  "..",
  "..",
  "..",
  "..",
  "crates",
  "text-shaper",
  "tests",
  "fixtures",
  "BidiCharacterTest-subset.txt",
);

/**
 * Runs a script in a fresh process, where no earlier test has loaded the
 * artifact yet and `fetch` refuses every request, so a load that reaches for
 * the network fails instead of passing on a cached instance.
 */
const runIsolated = async (body: string): Promise<string> => {
  const script = `
    globalThis.fetch = () => { throw new Error("fetch is not allowed here"); };
    const { getShaper } = await import(${JSON.stringify(SHAPER_MODULE)});
    const { readFileSync } = await import("node:fs");
    const wasm = readFileSync(${JSON.stringify(WASM_PATH)});
    ${body}
  `;
  const child = Bun.spawn([process.execPath, "--eval", script], {
    stdout: "pipe",
    stderr: "pipe",
  });
  const [stdout, stderr, code] = await Promise.all([
    new Response(child.stdout).text(),
    new Response(child.stderr).text(),
    child.exited,
  ]);
  if (code !== 0) {
    throw new Error(`isolated script failed (${code}): ${stderr}`);
  }
  return stdout.trim();
};

describe("loading from bytes", () => {
  test("instantiates the given bytes without fetching the artifact", async () => {
    const out = await runIsolated(`
      const shaper = await getShaper({ wasm });
      console.log(JSON.stringify(shaper.resolveBidi({ text: "a\\u05d0", direction: "ltr" })));
    `);
    expect(JSON.parse(out)).toEqual({ paragraphLevel: 0, levels: [0, 1], visualOrder: [0, 1] });
  });

  test("accepts an ArrayBuffer as well as a Uint8Array", async () => {
    const out = await runIsolated(`
      const buffer = wasm.buffer.slice(wasm.byteOffset, wasm.byteOffset + wasm.byteLength);
      const shaper = await getShaper({ wasm: buffer });
      console.log(shaper.resolveBidi({ text: "\\u05d0", direction: "auto" }).paragraphLevel);
    `);
    expect(out).toBe("1");
  });

  test("a failed load does not stop a later one", async () => {
    const out = await runIsolated(`
      const first = await getShaper({ wasm: new Uint8Array([0, 1, 2, 3]) }).then(
        () => "loaded",
        (error) => error._tag,
      );
      const second = await getShaper({ wasm });
      console.log(first, second.resolveBidi({ text: "a", direction: "ltr" }).levels.join());
    `);
    expect(out).toBe("ShaperError 0");
  });
});

type ConformanceCase = {
  readonly line: number;
  readonly text: string;
  readonly direction: BidiDirection;
  readonly paragraphLevel: number;
  /** `null` where rule X9 removes the character. */
  readonly levels: readonly (number | null)[];
  readonly visualOrder: readonly number[];
};

const DIRECTION_BY_FIELD: Record<string, BidiDirection> = {
  "0": BIDI_DIRECTION.leftToRight,
  "1": BIDI_DIRECTION.rightToLeft,
  "2": BIDI_DIRECTION.auto,
};

const numbersOf = (field: string): number[] => field.split(" ").map(Number);

const readConformanceCases = (): readonly ConformanceCase[] =>
  readFileSync(CONFORMANCE_FIXTURE, "utf8")
    .split("\n")
    .flatMap((raw, index) => {
      if (raw === "" || raw.startsWith("#")) {
        return [];
      }
      const [points = "", direction = "", paragraphLevel = "", levels = "", order = ""] =
        raw.split(";");
      const resolved = DIRECTION_BY_FIELD[direction];
      if (resolved === undefined) {
        throw new Error(`line ${index + 1}: unknown direction ${direction}`);
      }
      return [
        {
          line: index + 1,
          text: String.fromCodePoint(...points.split(" ").map((hex) => Number.parseInt(hex, 16))),
          direction: resolved,
          paragraphLevel: Number(paragraphLevel),
          levels: levels.split(" ").map((level) => (level === "x" ? null : Number(level))),
          visualOrder: numbersOf(order),
        },
      ];
    });

describe("the bidirectional algorithm", () => {
  test("resolves the Unicode conformance cases through the artifact", async () => {
    const shaper = await getShaper();
    const cases = readConformanceCases();
    expect(cases.length).toBeGreaterThan(500);
    const failures = cases.flatMap((expected) => {
      const line = shaper.resolveBidi({ text: expected.text, direction: expected.direction });
      const levels = line.levels.map((level, index) =>
        expected.levels[index] === null ? null : level,
      );
      const visualOrder = line.visualOrder.filter((index) => expected.levels[index] !== null);
      const matches =
        line.paragraphLevel === expected.paragraphLevel &&
        Bun.deepEquals(levels, expected.levels) &&
        Bun.deepEquals(visualOrder, expected.visualOrder);
      return matches ? [] : [expected.line];
    });
    expect(failures).toEqual([]);
  });

  test("counts code points, not UTF-16 units", async () => {
    const shaper = await getShaper();
    // An astral Adlam letter is two UTF-16 units and one code point.
    const line = shaper.resolveBidi({ text: "a\u{1e900}b", direction: BIDI_DIRECTION.leftToRight });
    expect(line.levels).toEqual([0, 1, 0]);
    expect(line.visualOrder).toEqual([0, 1, 2]);
  });

  test("an isolated right-to-left name leaves a following date where it was written", async () => {
    const shaper = await getShaper();
    const visual = (text: string): string => {
      const characters = [...text];
      return shaper
        .resolveBidi({ text, direction: BIDI_DIRECTION.leftToRight })
        .visualOrder.map((index) => characters[index] ?? "")
        .join("")
        .replaceAll(/[⁦-⁩]/gu, "");
    };
    const name = "محمد";
    const reversed = "دمحم";
    expect(visual(`by ${name} 2026`)).toBe(`by 2026 ${reversed}`);
    expect(visual(`by ⁨${name}⁩ 2026`)).toBe(`by ${reversed} 2026`);
  });
});
