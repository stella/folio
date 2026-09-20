/**
 * The survival census: run the survival law over every pair the schema declares.
 *
 * The census is the evidence the container contract is written against. It
 * walks the committed schema graph, synthesises one package per (container,
 * allowed child) and (element, allowed attribute) pair inside the parts folio
 * rebuilds, runs the four laws, and records every loss with the mechanism that
 * caused it. Its output is a shrink-only baseline: a pair that starts being
 * lost fails CI, and a pair that stops being lost fails CI too until the win is
 * locked in, so the list can only go down.
 *
 * Usage:
 *   bun scripts/container-survival-census.ts run
 *   bun scripts/container-survival-census.ts run --only tblGridChange
 *   bun scripts/container-survival-census.ts run --check
 *   bun scripts/container-survival-census.ts run --write-baseline
 *   bun scripts/container-survival-census.ts fixture --only tblGridChange
 *
 * `--only` matches a substring of the pair key, `--kind child|attribute`
 * narrows the space, and `--concurrency` bounds the parallelism (2 by default:
 * every pair writes and re-reads two packages, and the machine is shared).
 */

import path from "node:path";

import { TaggedError } from "better-result";

import { buildFixture, type Subject } from "./lib/container-survival/fixture";
import {
  bodyOf,
  editorSavePart,
  forcedSavePart,
  LOSS_MECHANISMS,
  type LossMechanism,
  type PairOutcome,
  probeOf,
  runSurvivalLaws,
  subjectKey,
  SURVIVAL_LAWS,
  type SurvivalLaw,
} from "./lib/container-survival/laws";
import {
  attributeSlots,
  childSlots,
  type ContainerSpace,
  loadContainerSpace,
} from "./lib/container-survival/schemaSpace";
import { representativeValue, valuesForType } from "./lib/container-survival/values";

const REPOSITORY_ROOT = path.resolve(import.meta.dir, "..");
const BASELINE_PATH = path.join(
  REPOSITORY_ROOT,
  "specifications/container-contract/survival-baseline.json",
);

class ContainerSurvivalError extends TaggedError("ContainerSurvivalError")<{
  message: string;
}> {}

export type SurvivalBaseline = {
  schemaVersion: 1;
  /** Every pair the sweep could build and run. */
  pairs: number;
  /** Pairs the fixture builder cannot represent, grouped by reason. */
  unrepresentable: Record<string, number>;
  /** The shrink-only list: pair key to the mechanism that loses it. */
  losses: Record<string, LossMechanism | "threw">;
  /**
   * Losses that only some values of a slot suffer, keyed `<pair key>=<value>`.
   *
   * A slot that survives its representative value can lose every other one: an
   * enumeration member nobody mapped, a `0` a truthiness test swallowed, the
   * Strict spelling of a measure. Kept apart from {@link losses} because the
   * contract decides slots, not values, and mixing the two would make the
   * contract's universe depend on which value the census happened to write.
   */
  valueLosses: Record<string, LossMechanism | "threw">;
};

const EMPTY_BASELINE: SurvivalBaseline = {
  schemaVersion: 1,
  pairs: 0,
  unrepresentable: {},
  losses: {},
  valueLosses: {},
};

/** Every pair in the space, child slots first, each with its representative value. */
export const allSubjects = (space: ContainerSpace): Subject[] => [
  ...childSlots(space).map((slot): Subject => ({ kind: "child", slot })),
  ...attributeSlots(space).flatMap((slot): Subject[] => {
    const value = representativeValue(space.index, slot.typeQName);
    return value === undefined ? [] : [{ kind: "attribute", slot, value }];
  }),
];

/** The other values an attribute's simple type accepts, beyond the representative one. */
export const otherValuesOf = (space: ContainerSpace, subject: Subject): string[] =>
  subject.kind === "attribute"
    ? valuesForType(space.index, subject.slot.typeQName).values.filter(
        (value) => value !== subject.value,
      )
    : [];

export const valueKey = (subject: Subject): string =>
  subject.kind === "attribute" ? `${subjectKey(subject)}=${subject.value}` : subjectKey(subject);

type Options = {
  only: string | undefined;
  kind: "child" | "attribute" | undefined;
  limit: number | undefined;
  concurrency: number;
  check: boolean;
  writeBaseline: boolean;
};

const readOptions = (argv: readonly string[]): Options => {
  const valueAfter = (flag: string): string | undefined => {
    const index = argv.indexOf(flag);
    return index === -1 ? undefined : argv[index + 1];
  };
  const kind = valueAfter("--kind");
  if (kind !== undefined && kind !== "child" && kind !== "attribute") {
    throw new ContainerSurvivalError({ message: `--kind takes child or attribute, not ${kind}` });
  }
  const limit = valueAfter("--limit");
  return {
    only: valueAfter("--only"),
    kind,
    limit: limit === undefined ? undefined : Number.parseInt(limit, 10),
    concurrency: Number.parseInt(valueAfter("--concurrency") ?? "2", 10),
    check: argv.includes("--check"),
    writeBaseline: argv.includes("--write-baseline"),
  };
};

const select = (space: ContainerSpace, options: Options): Subject[] => {
  let subjects = allSubjects(space);
  if (options.kind !== undefined) {
    subjects = subjects.filter((subject) => subject.kind === options.kind);
  }
  if (options.only !== undefined) {
    const only = options.only;
    subjects = subjects.filter((subject) => subjectKey(subject).includes(only));
  }
  return options.limit === undefined ? subjects : subjects.slice(0, options.limit);
};

/** A bounded worker pool; the census writes two packages per pair and the machine is shared. */
const mapBounded = async <T, R>(
  items: readonly T[],
  concurrency: number,
  run: (item: T) => Promise<R>,
): Promise<R[]> => {
  const results: R[] = Array.from({ length: items.length });
  let next = 0;
  const worker = async (): Promise<void> => {
    for (;;) {
      const index = next;
      next += 1;
      const item = items[index];
      if (item === undefined) {
        return;
      }
      results[index] = await run(item);
    }
  };
  await Promise.all(Array.from({ length: Math.max(1, concurrency) }, worker));
  return results;
};

const LAW_ORDER: readonly SurvivalLaw[] = [
  SURVIVAL_LAWS.parse,
  SURVIVAL_LAWS.serialize,
  SURVIVAL_LAWS.editor,
  SURVIVAL_LAWS.schema,
];

export type Census = {
  outcomes: readonly PairOutcome[];
  baseline: SurvivalBaseline;
};

const lossOf = (outcome: PairOutcome): LossMechanism | "threw" | undefined => {
  if (outcome.unrepresentable !== null) {
    return undefined;
  }
  if (outcome.laws[SURVIVAL_LAWS.parse] === false) {
    return "threw";
  }
  if (outcome.mechanism !== null) {
    return outcome.mechanism;
  }
  return outcome.laws[SURVIVAL_LAWS.schema] === false ? LOSS_MECHANISMS.respelled : undefined;
};

const sortedEntries = <T>(record: Record<string, T>): Record<string, T> =>
  Object.fromEntries(Object.entries(record).sort(([left], [right]) => left.localeCompare(right)));

export const censusOf = (
  outcomes: readonly PairOutcome[],
  valueOutcomes: ReadonlyArray<{ key: string; outcome: PairOutcome }> = [],
): Census => {
  const unrepresentable: Record<string, number> = {};
  const losses: Record<string, LossMechanism | "threw"> = {};
  let pairs = 0;
  for (const outcome of outcomes) {
    if (outcome.unrepresentable !== null) {
      unrepresentable[outcome.unrepresentable] =
        (unrepresentable[outcome.unrepresentable] ?? 0) + 1;
      continue;
    }
    pairs += 1;
    const loss = lossOf(outcome);
    if (loss !== undefined) {
      losses[outcome.key] = loss;
    }
  }
  const valueLosses: Record<string, LossMechanism | "threw"> = {};
  for (const { key, outcome } of valueOutcomes) {
    if (outcome.unrepresentable !== null) {
      continue;
    }
    const loss = lossOf(outcome);
    if (loss !== undefined) {
      valueLosses[key] = loss;
    }
  }
  return {
    outcomes,
    baseline: {
      schemaVersion: 1,
      pairs,
      unrepresentable: sortedEntries(unrepresentable),
      losses: sortedEntries(losses),
      valueLosses: sortedEntries(valueLosses),
    },
  };
};

const loadBaseline = async (): Promise<SurvivalBaseline> => {
  const file = Bun.file(BASELINE_PATH);
  return (await file.exists()) ? ((await file.json()) as SurvivalBaseline) : EMPTY_BASELINE;
};

/**
 * The ratchet.
 *
 * A pair that starts being lost fails. A pair the baseline lists that now
 * survives fails too, until the baseline is rewritten: a fix that does not lock
 * its own win in leaves room for the next change to undo it silently.
 */
export const compareToBaseline = (
  measured: SurvivalBaseline,
  baseline: SurvivalBaseline,
  scoped: boolean,
): string[] => {
  const problems: string[] = [];
  const compare = (
    measuredLosses: Record<string, LossMechanism | "threw">,
    recordedLosses: Record<string, LossMechanism | "threw">,
  ): void => {
    for (const [key, mechanism] of Object.entries(measuredLosses)) {
      const recorded = recordedLosses[key];
      if (recorded === undefined) {
        problems.push(`new loss: ${key} (${mechanism})`);
        continue;
      }
      if (recorded !== mechanism) {
        problems.push(`loss changed mechanism: ${key} (${recorded} -> ${mechanism})`);
      }
    }
    if (scoped) {
      return;
    }
    for (const key of Object.keys(recordedLosses)) {
      if (measuredLosses[key] === undefined) {
        problems.push(`fixed but not locked in: ${key}; re-run with --write-baseline`);
      }
    }
  };
  compare(measured.losses, baseline.losses);
  compare(measured.valueLosses, baseline.valueLosses);
  return problems;
};

const report = (census: Census): void => {
  const { outcomes, baseline } = census;
  const perLaw = new Map<SurvivalLaw, { held: number; broken: number; skipped: number }>(
    LAW_ORDER.map((law) => [law, { held: 0, broken: 0, skipped: 0 }]),
  );
  const byMechanism = new Map<string, number>();
  const byContainer = new Map<string, number>();

  for (const outcome of outcomes) {
    if (outcome.unrepresentable !== null) {
      continue;
    }
    for (const law of LAW_ORDER) {
      const tally = perLaw.get(law);
      const held = outcome.laws[law];
      if (tally === undefined) {
        continue;
      }
      if (held === null) {
        tally.skipped += 1;
      } else if (held) {
        tally.held += 1;
      } else {
        tally.broken += 1;
      }
    }
    const loss = lossOf(outcome);
    if (loss === undefined) {
      continue;
    }
    byMechanism.set(loss, (byMechanism.get(loss) ?? 0) + 1);
    // A pair lost because its container is says nothing about the container it
    // is declared in; charging it there would rank the wrong containers first.
    if (loss !== LOSS_MECHANISMS.containerLost) {
      byContainer.set(outcome.container, (byContainer.get(outcome.container) ?? 0) + 1);
    }
  }

  const unrepresentable = Object.values(baseline.unrepresentable).reduce(
    (total, count) => total + count,
    0,
  );
  console.log(
    `pairs generated ${outcomes.length}, run ${baseline.pairs}, unrepresentable ${unrepresentable}`,
  );
  for (const law of LAW_ORDER) {
    const tally = perLaw.get(law);
    console.log(
      `  ${law}: held ${tally?.held ?? 0}, broken ${tally?.broken ?? 0}, not reached ${tally?.skipped ?? 0}`,
    );
  }
  console.log(
    `losses ${Object.keys(baseline.losses).length}, value-only losses ${Object.keys(baseline.valueLosses).length}, by mechanism:`,
  );
  for (const [mechanism, count] of [...byMechanism].sort(([, a], [, b]) => b - a)) {
    console.log(`  ${count.toString().padStart(5)}  ${mechanism}`);
  }
  console.log("top containers by loss count, the container's own losses only:");
  for (const [container, count] of [...byContainer].sort(([, a], [, b]) => b - a).slice(0, 15)) {
    console.log(`  ${count.toString().padStart(5)}  ${container}`);
  }
  if (unrepresentable > 0) {
    console.log("unrepresentable, by reason:");
    for (const [reason, count] of Object.entries(baseline.unrepresentable).sort(
      ([, a], [, b]) => b - a,
    )) {
      console.log(`  ${count.toString().padStart(5)}  ${reason}`);
    }
  }
};

const runCensus = async (options: Options): Promise<number> => {
  const space = await loadContainerSpace();
  const subjects = select(space, options);
  const outcomes = await mapBounded(subjects, options.concurrency, (subject) =>
    runSurvivalLaws(space, subject),
  );

  // Only a slot whose representative value survived has anything left to say:
  // one already in `losses` would report the same loss once per value.
  const survivors = new Set(
    outcomes.filter((outcome) => lossOf(outcome) === undefined).map(({ key }) => key),
  );
  const valueSubjects: Subject[] = [];
  for (const subject of subjects) {
    if (subject.kind !== "attribute" || !survivors.has(subjectKey(subject))) {
      continue;
    }
    for (const value of otherValuesOf(space, subject)) {
      valueSubjects.push({ kind: "attribute", slot: subject.slot, value });
    }
  }
  const valueOutcomes = await mapBounded(valueSubjects, options.concurrency, async (subject) => ({
    key: valueKey(subject),
    outcome: await runSurvivalLaws(space, subject),
  }));

  const census = censusOf(outcomes, valueOutcomes);
  report(census);

  if (options.writeBaseline) {
    if (options.only !== undefined || options.limit !== undefined || options.kind !== undefined) {
      throw new ContainerSurvivalError({
        message: "a baseline may only be written from a full sweep; drop --only/--limit/--kind",
      });
    }
    await Bun.write(BASELINE_PATH, `${JSON.stringify(census.baseline, null, 2)}\n`);
    console.log(`wrote ${path.relative(REPOSITORY_ROOT, BASELINE_PATH)}`);
    return 0;
  }

  if (!options.check) {
    return 0;
  }
  const scoped =
    options.only !== undefined || options.limit !== undefined || options.kind !== undefined;
  const problems = compareToBaseline(census.baseline, await loadBaseline(), scoped);
  for (const problem of problems) {
    console.error(`  ${problem}`);
  }
  return problems.length === 0 ? 0 : 1;
};

/**
 * What a pair looked like going in and coming out of a forced save.
 *
 * The census says a pair is lost and names the mechanism; fixing it needs the
 * two strings side by side. Keeping that here rather than in a throwaway script
 * means the next person does not rewrite it.
 */
const explainPairs = async (options: Options): Promise<number> => {
  const space = await loadContainerSpace();
  for (const subject of select(space, { ...options, limit: options.limit ?? 5 })) {
    const built = buildFixture(space, subject);
    console.log(`--- ${subjectKey(subject).replaceAll(/\{[^}]*\}/gu, "")}`);
    if (built.status !== "built") {
      console.log(`    unrepresentable: ${built.reason}`);
      continue;
    }
    const outcome = await runSurvivalLaws(space, subject);
    const forced = await forcedSavePart(built.fixture);
    const probe = probeOf({ space, subject, fixture: built.fixture, xml: forced });
    console.log(`    mechanism: ${outcome.mechanism ?? "survives"}`);
    console.log(`    part: ${built.fixture.part.path}`);
    // Where the law looked and how many it wanted: a pair reported lost that
    // the markup below plainly contains is a pair found at another location.
    console.log(`    look: ${probe?.location ?? "?"}`);
    console.log(`    saw : ${probe?.found ?? 0} of ${probe?.expected ?? 0}`);
    console.log(`    in : ${bodyOf(built.fixture.documentXml)}`);
    console.log(`    out: ${bodyOf(forced)}`);
    if (outcome.mechanism === LOSS_MECHANISMS.editorProjection) {
      console.log(`    pm : ${bodyOf(await editorSavePart(built.fixture))}`);
    }
  }
  return 0;
};

const printFixtures = async (options: Options): Promise<number> => {
  const space = await loadContainerSpace();
  for (const subject of select(space, { ...options, limit: options.limit ?? 3 })) {
    const built = buildFixture(space, subject);
    console.log(`--- ${subjectKey(subject)}`);
    console.log(
      built.status === "built" ? built.fixture.documentXml : `unrepresentable: ${built.reason}`,
    );
  }
  return 0;
};

const main = (): Promise<number> => {
  const argv = Bun.argv.slice(2);
  const options = readOptions(argv);
  if (argv.includes("fixture")) {
    return printFixtures(options);
  }
  return argv.includes("explain") ? explainPairs(options) : runCensus(options);
};

if (import.meta.main) {
  process.exitCode = await main();
}
