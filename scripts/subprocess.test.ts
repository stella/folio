/** A test's child is judged by how it ended, and a misbehaving child is reported. */
import { expect, setDefaultTimeout, test } from "bun:test";

import { scanTestTimeout } from "../test/scan-timeout";
import {
  childHangTimeout,
  expectExited,
  runChild,
  subprocessTestTimeout,
} from "../test/subprocess";

setDefaultTimeout(subprocessTestTimeout());

// A child that never exits is hung however fast the runner is, so a short
// guard proves the hang report without making it depend on runner speed.
const NEVER_EXITS_GUARD = { hangTimeoutMs: scanTestTimeout(1_000) };

const bunEval = (source: string): string[] => [process.execPath, "-e", source];

test("a child that exits reports its status and output", () => {
  expect(
    runChild(
      bunEval('process.stdout.write("out"); process.stderr.write("err")'),
    ),
  ).toEqual({ kind: "exited", exitCode: 0, stdout: "out", stderr: "err" });
});

test("a child reads the input it is given", () => {
  expect(
    runChild(bunEval("process.stdout.write(await Bun.stdin.text())"), {
      input: "fed",
    }),
  ).toEqual({ kind: "exited", exitCode: 0, stdout: "fed", stderr: "" });
});

test("a child that exits non-zero is reported with its status", () => {
  expect(runChild(bunEval("process.exit(3)"))).toMatchObject({
    kind: "exited",
    exitCode: 3,
  });
});

test("a child that never exits is reported as hung with what it wrote", () => {
  expect(
    runChild(
      bunEval('process.stdout.write("started"); setInterval(() => {}, 1_000)'),
      NEVER_EXITS_GUARD,
    ),
  ).toEqual({ kind: "hung", stdout: "started", stderr: "" });
});

test("a child that loops without progress is reported as hung", () => {
  expect(runChild(bunEval("for (;;) {}"), NEVER_EXITS_GUARD)).toMatchObject({
    kind: "hung",
  });
});

test("a child ended by a signal is reported with the signal", () => {
  expect(
    runChild(bunEval('process.kill(process.pid, "SIGTERM")')),
  ).toMatchObject({ kind: "killed", signal: "SIGTERM" });
});

test("a child that must exit hands back its exit, and fails the test otherwise", () => {
  expect(expectExited(runChild(bunEval("process.exit(4)")))).toMatchObject({
    kind: "exited",
    exitCode: 4,
  });
  expect(() =>
    expectExited(
      runChild(
        bunEval(
          'process.stdout.write("started"); setInterval(() => {}, 1_000)',
        ),
        NEVER_EXITS_GUARD,
      ),
    ),
  ).toThrow(/The child hung\.\nstdout:\nstarted/u);
  expect(() =>
    expectExited(runChild(bunEval('process.kill(process.pid, "SIGTERM")'))),
  ).toThrow("The child was killed by SIGTERM.");
});

test("the test budget always outlasts its children's hang guards", () => {
  expect(subprocessTestTimeout()).toBeGreaterThan(childHangTimeout());
  expect(subprocessTestTimeout(3)).toBeGreaterThan(3 * childHangTimeout());
  expect(subprocessTestTimeout(1, 240_000)).toBe(
    childHangTimeout() + scanTestTimeout(240_000),
  );
  for (const children of [0, -1, 1.5, Number.NaN]) {
    expect(() => subprocessTestTimeout(children)).toThrow("whole number");
  }
});
