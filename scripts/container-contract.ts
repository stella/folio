/**
 * The container contract: write it from a census, then hold the code to it.
 *
 * `write` seeds the contract from a fresh survival census: every pair the
 * schema allows inside the parts folio rebuilds gets a disposition, and every
 * `dropped` one gets a reason class naming the mechanism and the fix. `check`
 * runs a fresh census against the committed contract and fails on either kind
 * of disagreement — a pair declared kept that is now lost, and a pair declared
 * dropped that now survives. The second direction is what stops the list
 * rotting: a fix has to move its pairs out of `dropped` explicitly.
 *
 * The keys are derived, never hand-listed. They come from
 * `scripts/lib/container-survival/schemaSpace.ts`, which walks the committed
 * schema graph, so a schema refresh that adds an element adds its pairs here
 * and `check` reports them as undeclared.
 *
 * Totality is enforced by this check rather than by a generated union and a
 * `Record<Union, ContractEntry>`. The union would hold about three thousand
 * members and the `satisfies` over it would cost more type instantiations than
 * every published package put together; the check costs a run of a census that
 * already has to happen. The compiler still owns what a decision may *say*:
 * `ContractEntry` is a discriminated union, and a `dropped` entry that names no
 * reason, or names one `DROP_REASONS` does not define, fails `typecheck`.
 *
 * Usage:
 *   bun scripts/container-contract.ts write
 *   bun scripts/container-contract.ts check
 *   bun scripts/container-contract.ts check --only tblGridChange
 */

import path from "node:path";

import { TaggedError } from "better-result";

import type {
  ContainerContract,
  ContractEntry,
  DropReason,
} from "../specifications/container-contract/dispositions";
import { DISPOSITIONS, DROP_REASONS } from "../specifications/container-contract/dispositions";
import { allSubjects } from "./container-survival-census";
import {
  LOSS_MECHANISMS,
  type LossMechanism,
  type PairOutcome,
  runSurvivalLaws,
  subjectKey,
  SURVIVAL_LAWS,
} from "./lib/container-survival/laws";
import { loadContainerSpace } from "./lib/container-survival/schemaSpace";

const REPOSITORY_ROOT = path.resolve(import.meta.dir, "..");
const CONTRACT_PATH = path.join(REPOSITORY_ROOT, "specifications/container-contract/contract.json");

class ContainerContractError extends TaggedError("ContainerContractError")<{
  message: string;
}> {}

/**
 * The reason class each loss mechanism belongs to.
 *
 * Total over the mechanisms by construction: a mechanism added to
 * `LOSS_MECHANISMS` without a class here fails `typecheck`, which is the point
 * — a new way to lose something must be given a name and a fix before it can
 * be recorded.
 */
const REASON_FOR_MECHANISM = {
  [LOSS_MECHANISMS.containerLost]: "containerNotKept",
  [LOSS_MECHANISMS.neverParsed]: "neverParsed",
  [LOSS_MECHANISMS.parsedNotSerialized]: "parsedNotSerialized",
  [LOSS_MECHANISMS.replayOnly]: "replayOnly",
  [LOSS_MECHANISMS.replayRejected]: "replayRejected",
  [LOSS_MECHANISMS.editorProjection]: "editorProjection",
  [LOSS_MECHANISMS.repeatTruncated]: "repeatTruncated",
  [LOSS_MECHANISMS.respelled]: "respelled",
} as const satisfies Record<LossMechanism, DropReason>;

const entryFor = (outcome: PairOutcome): ContractEntry | undefined => {
  if (outcome.unrepresentable !== null) {
    return undefined;
  }
  if (outcome.laws[SURVIVAL_LAWS.parse] === false) {
    return { disposition: DISPOSITIONS.dropped, reason: "parserThrows" };
  }
  if (outcome.mechanism !== null) {
    return { disposition: DISPOSITIONS.dropped, reason: REASON_FOR_MECHANISM[outcome.mechanism] };
  }
  return outcome.carrier === "capture"
    ? { disposition: DISPOSITIONS.capturedVerbatim }
    : { disposition: DISPOSITIONS.modelled };
};

const sameEntry = (left: ContractEntry, right: ContractEntry): boolean =>
  left.disposition === right.disposition &&
  (left.disposition !== DISPOSITIONS.dropped ||
    right.disposition !== DISPOSITIONS.dropped ||
    left.reason === right.reason);

const describe = (entry: ContractEntry): string =>
  entry.disposition === DISPOSITIONS.dropped
    ? `${entry.disposition} (${entry.reason})`
    : entry.disposition;

const loadContract = async (): Promise<ContainerContract> => {
  const file = Bun.file(CONTRACT_PATH);
  if (!(await file.exists())) {
    return { schemaVersion: 1, entries: {} };
  }
  return (await file.json()) as ContainerContract;
};

type Measured = { entries: Record<string, ContractEntry>; unrepresentable: number };

const measure = async (only: string | undefined, concurrency: number): Promise<Measured> => {
  const space = await loadContainerSpace();
  const subjects = allSubjects(space).filter(
    (subject) => only === undefined || subjectKey(subject).includes(only),
  );
  const entries: Record<string, ContractEntry> = {};
  let unrepresentable = 0;
  let next = 0;
  const worker = async (): Promise<void> => {
    for (;;) {
      const subject = subjects[next];
      next += 1;
      if (subject === undefined) {
        return;
      }
      const outcome = await runSurvivalLaws(space, subject);
      const entry = entryFor(outcome);
      if (entry === undefined) {
        unrepresentable += 1;
        continue;
      }
      entries[outcome.key] = entry;
    }
  };
  await Promise.all(Array.from({ length: Math.max(1, concurrency) }, worker));
  return { entries, unrepresentable };
};

const sortEntries = (entries: Record<string, ContractEntry>): Record<string, ContractEntry> =>
  Object.fromEntries(Object.entries(entries).sort(([left], [right]) => left.localeCompare(right)));

const report = (entries: Record<string, ContractEntry>, unrepresentable: number): void => {
  const counts = new Map<string, number>();
  for (const entry of Object.values(entries)) {
    const label = describe(entry);
    counts.set(label, (counts.get(label) ?? 0) + 1);
  }
  console.log(`${Object.keys(entries).length} pairs decided, ${unrepresentable} unrepresentable`);
  for (const [label, count] of [...counts].sort(([, a], [, b]) => b - a)) {
    console.log(`  ${String(count).padStart(5)}  ${label}`);
  }
};

const write = async (concurrency: number): Promise<number> => {
  const { entries, unrepresentable } = await measure(undefined, concurrency);
  report(entries, unrepresentable);
  const contract: ContainerContract = { schemaVersion: 1, entries: sortEntries(entries) };
  await Bun.write(CONTRACT_PATH, `${JSON.stringify(contract, null, 2)}\n`);
  console.log(`wrote ${path.relative(REPOSITORY_ROOT, CONTRACT_PATH)}`);
  return 0;
};

const check = async (only: string | undefined, concurrency: number): Promise<number> => {
  const contract = await loadContract();
  const { entries, unrepresentable } = await measure(only, concurrency);
  report(entries, unrepresentable);

  const problems: string[] = [];
  for (const [key, measured] of Object.entries(entries)) {
    const declared = contract.entries[key];
    if (declared === undefined) {
      problems.push(`no decision recorded: ${key} is ${describe(measured)}`);
      continue;
    }
    if (!sameEntry(declared, measured)) {
      problems.push(
        `contract says ${describe(declared)}, folio does ${describe(measured)}: ${key}`,
      );
    }
  }
  if (only === undefined) {
    for (const key of Object.keys(contract.entries)) {
      if (entries[key] === undefined) {
        problems.push(
          `decision for a pair the schema no longer declares, or no longer runs: ${key}`,
        );
      }
    }
  }

  for (const problem of problems) {
    console.error(`  ${problem}`);
  }
  if (problems.length > 0) {
    console.error(
      "The contract and the code disagree. Fix the code, or record the new decision with `bun run container-contract:write`.",
    );
  }
  return problems.length === 0 ? 0 : 1;
};

const main = (): Promise<number> => {
  const argv = Bun.argv.slice(2);
  const valueAfter = (flag: string): string | undefined => {
    const index = argv.indexOf(flag);
    return index === -1 ? undefined : argv[index + 1];
  };
  const concurrency = Number.parseInt(valueAfter("--concurrency") ?? "2", 10);
  if (argv.includes("write")) {
    return write(concurrency);
  }
  if (argv.includes("check")) {
    return check(valueAfter("--only"), concurrency);
  }
  throw new ContainerContractError({ message: "expected `write` or `check`" });
};

// Referenced so a reason class that nothing maps to still has to exist.
void DROP_REASONS;

if (import.meta.main) {
  process.exitCode = await main();
}
