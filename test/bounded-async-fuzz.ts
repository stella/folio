import { panic } from "better-result";
import fc from "fast-check";

export const BROWSER_FUZZ_BUDGET = {
  testMs: 600_000,
  discoveryMs: 360_000,
  shrinkMs: 60_000,
} as const;

/**
 * The legacy editor's seeds share the canonical lane's 35-minute step, so each
 * has a smaller cap. A finding ends discovery early, and 20 passing traces took
 * about 41 seconds in that step, which leaves this cap well above a healthy seed.
 */
export const DEFAULT_EDITOR_FUZZ_BUDGET = {
  discoveryMs: 150_000,
  shrinkMs: 30_000,
} as const;

type FirstFuzzFailure<T> = {
  seed: number;
  path: string;
  value: T;
  error: unknown;
};

type BoundedAsyncFuzzOptions<T> = {
  arbitrary: fc.Arbitrary<T>;
  evaluate: (value: T) => Promise<void>;
  seed: number;
  numRuns: number;
  path?: string;
  discoveryMs: number;
  shrinkMs: number;
  onFirstFailure: (failure: FirstFuzzFailure<T>) => Promise<void>;
};

/** Retain the first witness before giving its shrinker a separate time budget. */
export const checkWithBoundedShrink = async <T>({
  arbitrary,
  evaluate,
  seed,
  numRuns,
  path,
  discoveryMs,
  shrinkMs,
  onFirstFailure,
}: BoundedAsyncFuzzOptions<T>) => {
  const discovery = await fc.check(fc.asyncProperty(arbitrary, evaluate), {
    seed,
    numRuns,
    ...(path === undefined ? {} : { path }),
    endOnFailure: true,
    interruptAfterTimeLimit: discoveryMs,
    markInterruptAsFailure: true,
  });
  if (
    !discovery.failed ||
    discovery.counterexample === null ||
    discovery.counterexamplePath === null
  )
    return { verdict: discovery, discovery, shrinking: null };
  const value = discovery.counterexample[0];
  await onFirstFailure({
    seed: discovery.seed,
    path: discovery.counterexamplePath,
    value,
    error: discovery.errorInstance,
  });
  const replay: { status: "pending" | "matched" | "mismatch" } = {
    status: "pending",
  };
  const shrinking = await fc.check(
    fc.asyncProperty(arbitrary, (candidate) => {
      if (replay.status !== "pending") return evaluate(candidate);
      if (fc.stringify(candidate) !== fc.stringify(value)) {
        replay.status = "mismatch";
        throw new fc.PreconditionFailure(true);
      }
      replay.status = "matched";
      // The first witness is already known to fail; never rerun it until green.
      throw discovery.errorInstance;
    }),
    {
      seed: discovery.seed,
      path: discovery.counterexamplePath,
      numRuns: 1,
      endOnFailure: false,
      interruptAfterTimeLimit: shrinkMs,
      markInterruptAsFailure: true,
    },
  );
  if (replay.status === "mismatch")
    panic("The seeded shrink replay did not start from the retained failure.");
  const verdict =
    shrinking.failed && !shrinking.interrupted ? shrinking : discovery;
  return { verdict, discovery, shrinking };
};
