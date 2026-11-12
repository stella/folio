/**
 * Child processes in tests. A test states what its child must do (exit
 * status and output); the only clock is a hang guard taken from the shared
 * scaled budget, far above any healthy run, so a slow runner never fails a
 * child that would have finished. A child that never exits is reported as
 * hung by its own test instead of the runner's generic test timeout.
 */
import { panic } from "better-result";

import { scanTestTimeout } from "./scan-timeout";

const CHILD_HANG_BASE_MS = 60_000;
const TEST_SETUP_BASE_MS = 10_000;

/** How long one child may run before it counts as hung. */
export const childHangTimeout = (): number =>
  scanTestTimeout(CHILD_HANG_BASE_MS);

/**
 * Budget for a test that runs `children` children one after another: each
 * child's hang guard plus the test's own setup, so the child's hang report
 * always arrives before the test runner gives up on the test.
 */
export const subprocessTestTimeout = (children = 1): number => {
  if (!Number.isSafeInteger(children) || children < 1) {
    return panic(
      "A subprocess test runs a positive whole number of children.",
      {
        children,
      },
    );
  }
  return childHangTimeout() * children + scanTestTimeout(TEST_SETUP_BASE_MS);
};

type ChildOutput = {
  readonly stdout: string;
  readonly stderr: string;
};

export type ChildOutcome =
  | (ChildOutput & { readonly kind: "exited"; readonly exitCode: number })
  | (ChildOutput & { readonly kind: "hung" })
  | (ChildOutput & { readonly kind: "killed"; readonly signal: string });

export type ChildOptions = {
  readonly cwd?: string | undefined;
  readonly env?: Readonly<Record<string, string | undefined>> | undefined;
  /**
   * Only for proving the hang report itself with a child that never exits;
   * take it from scanTestTimeout. Every other caller keeps the default.
   */
  readonly hangTimeoutMs?: number | undefined;
};

const decoder = new TextDecoder();

/** Run a child to completion and report how it ended. */
export const runChild = (
  command: readonly string[],
  options: ChildOptions = {},
): ChildOutcome => {
  const result = Bun.spawnSync([...command], {
    cwd: options.cwd,
    env: options.env === undefined ? undefined : { ...options.env },
    stdin: "ignore",
    stdout: "pipe",
    stderr: "pipe",
    timeout: options.hangTimeoutMs ?? childHangTimeout(),
    killSignal: "SIGKILL",
  });
  const output = {
    stdout: decoder.decode(result.stdout),
    stderr: decoder.decode(result.stderr),
  };
  if (result.exitedDueToTimeout === true) return { kind: "hung", ...output };
  // Bun types exitCode as a number; it is null when a signal ended the child.
  const exitCode: number | null = result.exitCode;
  if (exitCode === null) {
    return {
      kind: "killed",
      signal: String(result.signalCode ?? "unknown"),
      ...output,
    };
  }
  return { kind: "exited", exitCode, ...output };
};
