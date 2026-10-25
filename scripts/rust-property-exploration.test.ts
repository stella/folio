import { describe, expect, test } from "bun:test";

import { failureRecord } from "../test/consumer-scenarios/support/failure-fingerprints";
import { extractFailureMarkers } from "./failure-fingerprints";
import { parseFailureRecord } from "./fuzz-failure-issues";
import {
  classifyRun,
  explorationSeed,
  findingMarker,
  parseCargoTestOutput,
  propertyTargets,
} from "./rust-property-exploration";

const lib = `     Running unittests src/lib.rs (target/debug/deps/stella_docx_kernel-1a2b)

running 2 tests
test semantic::tests::split_runs ... ok
test semantic::tests::prefixes ... ok

test result: ok. 2 passed; 0 failed; 0 ignored; 0 measured; 0 filtered out; finished in 0.10s
`;

const failing = (
  name: string,
  output: string,
) => `     Running tests/property_styles.rs (target/debug/deps/property_styles-3c4d)

running 2 tests
test ${name} ... FAILED
test plain ... ok

failures:

---- ${name} stdout ----
${output}
note: run with \`RUST_BACKTRACE=1\` environment variable to display a backtrace


failures:
    ${name}

test result: FAILED. 1 passed; 1 failed; 0 ignored; 0 measured; 0 filtered out; finished in 0.20s

error: test failed, to rerun pass \`-p stella-docx-kernel --test property_styles\`
`;

const generatedFailure = `proptest: Saving this and future failures in /w/crates/docx-kernel/tests/property_styles.proptest-regressions
proptest: If this test was run on a CI system, you may wish to add the following line to your copy of the file.
cc 0123abcd # shrinks to toggles = [true]

thread 'toggles_cancel' panicked at crates/docx-kernel/tests/property_styles.rs:91:1:
Test failed: assertion failed: \`(left == right)\`
  left: \`1\`,
 right: \`0\` at crates/docx-kernel/tests/property_styles.rs:120.
minimal failing input: toggles = [
    true,
]
	successes: 3
	local rejects: 0
	global rejects: 0
`;

const OPTIONS = {
  package: "stella-docx-kernel",
  crate: "crates/docx-kernel",
  seed: 4_000_000_001,
};

describe("property targets", () => {
  test("the library and each integration file with a property", () => {
    expect(
      propertyTargets([
        { path: "src/semantic.rs", text: "proptest! { #[test] fn a() {} }" },
        { path: "src/lib.rs", text: "pub mod semantic;" },
        { path: "tests/b.rs", text: "proptest::proptest! {}" },
        { path: "tests/a.rs", text: "proptest! {}" },
        { path: "tests/plain.rs", text: "#[test] fn plain() {}" },
        { path: "tests/common/mod.rs", text: "proptest! {}" },
      ]),
    ).toEqual(["--lib", "--test", "a", "--test", "b"]);
    expect(propertyTargets([{ path: "tests/a.rs", text: "" }])).toEqual([]);
  });

  test("the seed follows the commit and fits proptest's u64", () => {
    const seed = explorationSeed("a".repeat(40));
    expect(explorationSeed("a".repeat(40))).toBe(seed);
    expect(explorationSeed("b".repeat(40))).not.toBe(seed);
    expect(Number.isSafeInteger(seed) && seed >= 0 && seed < 2 ** 32).toBe(
      true,
    );
  });
});

describe("classifying a run", () => {
  test("a generated case's failure is a finding; the run passes", () => {
    const binaries = parseCargoTestOutput(
      `${lib}\n${failing("toggles_cancel", generatedFailure)}`,
    );
    expect(binaries.map(({ source, result }) => [source, result])).toEqual([
      ["src/lib.rs", "ok"],
      ["tests/property_styles.rs", "failed"],
    ]);
    const { findings, blocking } = classifyRun(binaries, 101);
    expect(blocking).toEqual([]);
    expect(findings).toEqual([
      {
        target: ["--test", "property_styles"],
        source: "tests/property_styles.rs",
        test: "toggles_cancel",
        reason:
          "assertion failed: `(left == right)`\n  left: `1`,\n right: `0` at crates/docx-kernel/tests/property_styles.rs:120",
        pin: "cc 0123abcd # shrinks to toggles = [true]",
      },
    ]);
  });

  test("a panic outside a property, an abort, a crash or a build error fails", () => {
    const panic = `thread 'plain' panicked at tests/property_styles.rs:3:5:\nassertion failed: false\n`;
    expect(
      classifyRun(parseCargoTestOutput(failing("plain", panic)), 101).blocking,
    ).toEqual([
      "tests/property_styles.rs: plain failed outside a generated case",
    ]);
    const abort = `thread 'toggles_cancel' panicked at tests/property_styles.rs:91:1:\nTest aborted: Too many global rejects\n`;
    expect(
      classifyRun(parseCargoTestOutput(failing("toggles_cancel", abort)), 101)
        .blocking,
    ).toHaveLength(1);
    const crashed = `     Running tests/property_styles.rs (target/debug/deps/property_styles-3c4d)\n\nrunning 2 tests\nerror: test failed, to rerun pass \`--test property_styles\`\n\nCaused by:\n  process didn't exit successfully (signal: 11, SIGSEGV)\n`;
    expect(classifyRun(parseCargoTestOutput(crashed), 101).blocking).toEqual([
      "tests/property_styles.rs: no test result (crashed or aborted)",
    ]);
    expect(
      classifyRun(parseCargoTestOutput("error[E0425]: cannot find value"), 101)
        .blocking,
    ).toEqual([
      "cargo test exited 101 without a failing test",
      "cargo test ran no test binary",
    ]);
    // A mixed binary: the finding is kept, and the other failure still fails.
    const mixed = failing("toggles_cancel", generatedFailure).replace(
      "failures:\n    toggles_cancel\n",
      "failures:\n    toggles_cancel\n    unlisted\n",
    );
    const result = classifyRun(parseCargoTestOutput(mixed), 101);
    expect(result.findings).toHaveLength(1);
    expect(result.blocking).toEqual([
      "tests/property_styles.rs: unlisted failed outside a generated case",
    ]);
  });
});

test("a finding's marker and record reach the filer marked exploration", () => {
  const [binary] = parseCargoTestOutput(
    failing("toggles_cancel", generatedFailure),
  );
  if (binary === undefined) throw new Error("missing binary");
  const [finding] = classifyRun([binary], 101).findings;
  if (finding === undefined) throw new Error("missing finding");
  const marker = findingMarker(finding, OPTIONS);
  expect(marker).toMatchObject({
    test: "crates/docx-kernel/tests/property_styles.rs::toggles_cancel",
    seed: OPTIONS.seed,
    path: null,
    repro:
      "PROPTEST_RNG_SEED=4000000001 cargo test -p stella-docx-kernel --test property_styles -- 'toggles_cancel' --exact",
    exploration: true,
  });
  expect(
    extractFailureMarkers(`FOLIO_FAILURE ${JSON.stringify(marker)}`),
  ).toEqual([marker]);
  const record = parseFailureRecord(
    JSON.parse(
      JSON.stringify(failureRecord(marker, new Error(finding.reason))),
    ),
  );
  expect(record?.marker.exploration).toBe(true);
  // The same failure under another commit's seed files under one signature.
  expect(findingMarker(finding, { ...OPTIONS, seed: 7 }).fingerprint).toBe(
    marker.fingerprint,
  );
});
