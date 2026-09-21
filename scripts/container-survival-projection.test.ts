/**
 * The survival law's L3 leg projects with reuse declined, and stays that way.
 *
 * L3 measures whether the ProseMirror projection can carry a pair. It measures
 * that only while `fromProseDoc` rebuilds every record: once the conversion may
 * merge an untouched record back from the base document by reference, a pair
 * the projection drops returns through the base object and the law reports a
 * survival the editor never performed. The contract's ratchet would then lock
 * that in for every pair the census records as `editorProjection`.
 *
 * The conversion exposes a reuse choice, and this binding keeps the law from
 * drifting if that contract changes. It has two halves:
 *
 * 1. the law reaches the conversion only through `projectWithoutReuse`, so
 *    there is a single place the option has to be passed;
 * 2. that call passes `reuse: "none"` exactly when the conversion's own
 *    parameter list offers the choice — so adding the option without wiring the
 *    law fails here, and so does wiring the law against a conversion that
 *    dropped it again.
 *
 * A `fromProseDoc.length` check cannot state (2): the option arrives as a
 * defaulted third parameter, which leaves the declared arity at 2.
 */

import { describe, expect, test } from "bun:test";
import { readdirSync, readFileSync } from "node:fs";
import path from "node:path";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const LAW_DIR = path.join(REPO_ROOT, "scripts", "lib", "container-survival");
const HELPER_FILE = "projection.ts";
const CONVERSION_PATH = path.join(
  REPO_ROOT,
  "packages",
  "core",
  "src",
  "prosemirror",
  "conversion",
  "fromProseDoc.ts",
);
const CONVERSION_DECLARATION = "export function fromProseDoc(";

/** An import of the conversion or a call to it, as against a mention in prose. */
const REACHES_CONVERSION = /fromProseDoc\(|conversion\/fromProseDoc"/u;

const read = (file: string): string => readFileSync(file, "utf8");

/**
 * The text between the parentheses opened at `opening`.
 *
 * A parameter list carries parentheses of its own as soon as a type names a
 * function, so the scan balances rather than stopping at the first `)`.
 */
const balancedAt = (source: string, opening: number): string => {
  let depth = 0;
  for (let index = opening; index < source.length; index += 1) {
    const character = source[index];
    if (character === "(") {
      depth += 1;
      continue;
    }
    if (character !== ")") {
      continue;
    }
    depth -= 1;
    if (depth === 0) {
      return source.slice(opening + 1, index);
    }
  }
  throw new Error("the parenthesis opened at this point is never closed");
};

/** The parameters `fromProseDoc` declares, as its owning module writes them. */
const conversionParameters = (): string => {
  const source = read(CONVERSION_PATH);
  const declaration = source.indexOf(CONVERSION_DECLARATION);
  if (declaration === -1) {
    throw new Error(
      `${CONVERSION_PATH} no longer declares \`${CONVERSION_DECLARATION}\`, so this binding cannot read the option it guards`,
    );
  }
  return balancedAt(source, declaration + CONVERSION_DECLARATION.length - 1);
};

/** The arguments the law's helper hands the conversion. */
const helperArguments = (): string => {
  const source = read(path.join(LAW_DIR, HELPER_FILE));
  const call = source.lastIndexOf("fromProseDoc(");
  if (call === -1) {
    throw new Error(`${HELPER_FILE} no longer calls the conversion it exists to own`);
  }
  return balancedAt(source, call + "fromProseDoc".length);
};

describe("the survival law projects through one call", () => {
  test("no law module but the helper reaches the conversion", () => {
    const reaching = readdirSync(LAW_DIR)
      .filter((entry) => entry.endsWith(".ts") && entry !== HELPER_FILE)
      .filter((entry) => REACHES_CONVERSION.test(read(path.join(LAW_DIR, entry))));

    expect(reaching).toEqual([]);
  });

  test("the law's editor leg projects through the helper", () => {
    expect(read(path.join(LAW_DIR, "laws.ts"))).toContain("projectWithoutReuse(toProseDoc(");
  });
});

describe("the helper declines reuse as soon as there is reuse to decline", () => {
  test("it passes the option exactly when the conversion accepts one", () => {
    const offered = /\breuse\b/u.test(conversionParameters());

    expect({ passesReuseNone: /\breuse:\s*"none"/u.test(helperArguments()) }).toEqual({
      passesReuseNone: offered,
    });
  });
});
