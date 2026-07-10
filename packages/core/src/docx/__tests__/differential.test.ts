/**
 * Differential parser testing (folio vs external OOXML references).
 *
 * Projects folio's parse of every corpus fixture into a structural shape and
 * asserts it matches the same projection taken from python-docx and the Open XML
 * SDK — locking in parse parity across the whole suite, not just a single smoke
 * fixture.
 *
 * Reference parsers are optional host dependencies:
 *  - python-docx (`pip install python-docx`)
 *  - Open XML SDK (`dotnet build packages/core/scripts/differential/dotnet -c Release`)
 *
 * When a reference is missing the suite SKIPS locally so `bun test` stays
 * runnable without Python or .NET; CI installs both and sets
 * `DIFFERENTIAL_REQUIRED=1`, which turns a missing dependency into a failure
 * so the parity gate cannot silently pass. See
 * `packages/core/scripts/differential/README.md`.
 */

import { describe, expect, test } from "bun:test";
import { Glob } from "bun";
import path from "node:path";

import {
  DIFFERENTIAL_REFERENCES,
  isOpenXmlSdkAvailable,
  isPythonDocxAvailable,
  runDifferential,
  type DifferentialReference,
} from "../../../scripts/differential/diff";

const REQUIRED = process.env.DIFFERENTIAL_REQUIRED === "1";
const PYTHON_DOCX_TIMEOUT_MS = 5_000;
const OPEN_XML_SDK_TIMEOUT_MS = 15_000;

// The corpus: the parser fixtures plus the visual fixtures (repo root).
const FIXTURE_DIRS = [
  path.join(import.meta.dir, "__fixtures__"),
  path.resolve(import.meta.dir, "../../../../../tests/visual/fixtures"),
];

const fixtures = FIXTURE_DIRS.flatMap((dir) => [
  ...new Glob("**/*.docx").scanSync({ cwd: dir, absolute: true }),
]).sort();

const referenceAvailability: Record<DifferentialReference, () => boolean> = {
  "python-docx": isPythonDocxAvailable,
  "open-xml-sdk": isOpenXmlSdkAvailable,
};

const referenceSetupHints: Record<DifferentialReference, string> = {
  "python-docx": "pip install python-docx",
  "open-xml-sdk": "dotnet build packages/core/scripts/differential/dotnet -c Release",
};

for (const reference of DIFFERENTIAL_REFERENCES) {
  const isAvailable = referenceAvailability[reference];
  const fixtureTimeoutMs =
    reference === "open-xml-sdk" ? OPEN_XML_SDK_TIMEOUT_MS : PYTHON_DOCX_TIMEOUT_MS;

  describe(`differential parser harness (folio vs ${reference})`, () => {
    if (!isAvailable()) {
      const message = `${reference} not available; see scripts/differential/README.md (${referenceSetupHints[reference]})`;
      if (REQUIRED) {
        test(`${reference} is required when DIFFERENTIAL_REQUIRED=1`, () => {
          throw new Error(message);
        });
      } else {
        test.skip(message, () => {});
      }
      return;
    }

    test("corpus is non-empty", () => {
      expect(fixtures.length).toBeGreaterThan(0);
    });

    for (const fixture of fixtures) {
      const name = path.basename(fixture);
      test(
        `structural projection matches ${reference}: ${name}`,
        async () => {
          const result = await runDifferential(fixture, { reference });
          if (!result.ok) {
            if (result.reason === "infra") {
              throw new Error(`harness infrastructure failure: ${result.message}`);
            }
            throw new Error(
              `unexpected divergence on ${name} (${reference}):\n${JSON.stringify(result.divergences, null, 2)}\n\nfolio: ${JSON.stringify(result.folio, null, 2)}\nreference: ${JSON.stringify(result.reference, null, 2)}`,
            );
          }
          expect(result.ok).toBe(true);
        },
        fixtureTimeoutMs,
      );
    }
  });
}
