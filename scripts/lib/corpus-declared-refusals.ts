import { TaggedError } from "better-result";
import {
  DECLARED_REFUSAL_INVARIANTS,
  DECLARED_REFUSAL_REASONS,
  type CorpusDeclaredRefusal,
  type DeclaredRefusalInvariant,
  type DeclaredRefusalReason,
} from "./corpus-invariants/contract";

export type DeclaredRefusalCounts = Record<
  DeclaredRefusalInvariant,
  Record<DeclaredRefusalReason, number>
>;

export class DeclaredRefusalAccountingError extends TaggedError("DeclaredRefusalAccountingError")<{
  message: string;
}> {}

export const emptyDeclaredRefusalCounts = (): DeclaredRefusalCounts => ({
  [DECLARED_REFUSAL_INVARIANTS.opInverse]: { [DECLARED_REFUSAL_REASONS.SIGNED_PACKAGE]: 0 },
  [DECLARED_REFUSAL_INVARIANTS.opLocality]: { [DECLARED_REFUSAL_REASONS.SIGNED_PACKAGE]: 0 },
});

const INVARIANTS = Object.values(DECLARED_REFUSAL_INVARIANTS);
const REASONS = Object.values(DECLARED_REFUSAL_REASONS);

/** Refuse old or unknown accounting instead of merging it as zero. */
export const assertDeclaredRefusalCounts = (counts: DeclaredRefusalCounts): void => {
  if (counts === undefined || Object.keys(counts).length !== INVARIANTS.length)
    throw new DeclaredRefusalAccountingError({
      message: "Census lacks complete declared-refusal accounting; measure it again",
    });
  for (const invariant of INVARIANTS) {
    const reasons = counts[invariant];
    if (reasons === undefined || Object.keys(reasons).length !== REASONS.length)
      throw new DeclaredRefusalAccountingError({
        message: "Unknown or missing declared-refusal reasons",
      });
    for (const reason of REASONS) {
      if (!Number.isSafeInteger(reasons[reason]) || reasons[reason] < 0)
        throw new DeclaredRefusalAccountingError({ message: "Invalid declared-refusal count" });
    }
  }
};

/** A file contributes once per invariant/reason, even if its payload repeats it. */
export const countDeclaredRefusals = (
  counts: DeclaredRefusalCounts,
  refusals: readonly CorpusDeclaredRefusal[],
): void => {
  const seen = new Set<string>();
  for (const { invariant, reason } of refusals) {
    const key = `${invariant}:${reason}`;
    if (seen.has(key)) continue;
    seen.add(key);
    if (counts[invariant]?.[reason] === undefined)
      throw new DeclaredRefusalAccountingError({ message: "Unknown declared refusal" });
    counts[invariant][reason] += 1;
  }
};

export const mergeDeclaredRefusalCounts = (
  into: DeclaredRefusalCounts,
  from: DeclaredRefusalCounts,
): void => {
  assertDeclaredRefusalCounts(from);
  for (const invariant of INVARIANTS)
    for (const reason of REASONS) into[invariant][reason] += from[invariant][reason];
};

export const declaredRefusalCount = (counts: DeclaredRefusalCounts, family: string): number =>
  INVARIANTS.filter((invariant) => invariant === family).reduce(
    (total, invariant) =>
      total + REASONS.reduce((sum, reason) => sum + counts[invariant][reason], 0),
    0,
  );

export const renderDeclaredRefusals = (counts: DeclaredRefusalCounts): string[] =>
  INVARIANTS.flatMap((invariant) =>
    REASONS.map(
      (reason) => `  declared refusal ${invariant} / ${reason}: ${counts[invariant][reason]}`,
    ),
  );

export const declaredRefusalDifferences = (
  recorded: DeclaredRefusalCounts,
  observed: DeclaredRefusalCounts,
) => {
  assertDeclaredRefusalCounts(recorded);
  assertDeclaredRefusalCounts(observed);
  return INVARIANTS.flatMap((invariant) =>
    REASONS.flatMap((reason) => {
      const before = recorded[invariant][reason];
      const after = observed[invariant][reason];
      return before === after ? [] : [{ invariant, reason, before, after }];
    }),
  );
};
