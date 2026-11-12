import fc from "fast-check";

/**
 * What a property reaches, checked on a fixed sample: the same cases in every
 * environment, on every commit and run.
 *
 * A property's generated pass draws what its seed chooses, and may stop after
 * any case: on a failure, or on an exploration finding that pull request CI
 * reports without failing (test/property-exploration.ts). A tally written
 * inside the predicate therefore says nothing certain about which kinds the
 * generator reaches, and an assertion over it fails on seeds that never drew
 * a kind and on runs that ended early. Coverage a property must reach (every
 * kind applied, enough cases past the refusals) is checked on a fixed sample
 * instead, through the same predicate: the property keeps its fresh seed per
 * commit, and the coverage check fails only when the generator or the code
 * under test changes.
 *
 * scripts/property-draw-coverage.test.ts rejects assertions over tallies a
 * property's draw writes.
 */
const COVERAGE_SEED = 20_261_112;

export const coverageSample = <T>(
  arbitrary: fc.Arbitrary<T>,
  numRuns: number,
  seed: number = COVERAGE_SEED,
): T[] => fc.sample(arbitrary, { seed, numRuns });

type CoverageOptions<Ts> = {
  seed?: number;
  /** Rewrites the `index`th case, e.g. so that every kind takes its turn. */
  map?: (value: Ts, index: number) => Ts;
};

const settle = <Ts>(
  value: Ts,
  outcome: fc.PreconditionFailure | fc.PropertyFailure | null,
): void => {
  // A case its preconditions skip proves nothing, as in the property.
  if (outcome === null || outcome instanceof fc.PreconditionFailure) return;
  throw new Error(
    `The fixed coverage sample failed on ${fc.stringify(value)}`,
    { cause: outcome.error },
  );
};

/**
 * Run `property`'s predicate on a fixed sample of its own arbitraries, with
 * its hooks, failing on the first case that fails. Pass the property a tally
 * of its own and assert coverage over that tally afterwards.
 */
export function checkCoverageSample<Ts>(
  property: fc.IAsyncProperty<Ts>,
  numRuns: number,
  options?: CoverageOptions<Ts>,
): Promise<void>;
export function checkCoverageSample<Ts>(
  property: fc.IProperty<Ts>,
  numRuns: number,
  options?: CoverageOptions<Ts>,
): void;
export function checkCoverageSample<Ts>(
  property: fc.IRawProperty<Ts>,
  numRuns: number,
  { seed = COVERAGE_SEED, map = (value) => value }: CoverageOptions<Ts> = {},
): Promise<void> | void {
  const values = fc
    .sample(property, { seed, numRuns })
    .map((value, index) => map(value, index));
  if (property.isAsync()) {
    return (async () => {
      for (const value of values) {
        await property.runBeforeEach();
        const outcome = await property.run(value);
        await property.runAfterEach();
        settle(value, outcome);
      }
    })();
  }
  for (const value of values) {
    void property.runBeforeEach();
    const outcome = property.run(value) as
      fc.PreconditionFailure | fc.PropertyFailure | null;
    void property.runAfterEach();
    settle(value, outcome);
  }
}
