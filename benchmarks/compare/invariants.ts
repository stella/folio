/**
 * What every benchmark pair must satisfy before its timings mean anything.
 *
 * A faster comparison that produces a different document is not faster, it is
 * broken; a slower one that produces the same bytes is a real regression. So
 * each configuration carries both: the checks below, and the digests that let
 * a later run prove an optimization changed nothing.
 */

import { compareDocx } from "@stll/folio-core";
import { FolioDocxReviewer } from "@stll/folio-core/server";
import type { CompareChange, CompareDocxOptions } from "@stll/folio-core/compare/types";

import type { PackageValidator } from "./validator";

export const INVARIANTS = Object.freeze([
  /** Rejecting every generated revision returns the base. */
  "reject-returns-base",
  /** Accepting every generated revision returns the target. */
  "accept-returns-target",
  /** Comparing a document with itself invents nothing. */
  "self-compare-is-empty",
  /** Two runs over the same inputs produce the same bytes. */
  "byte-determinism",
  /** The redlined package validates against the OOXML schema. */
  "schema-validity",
  /**
   * A pair the harness made different is reported as different. The other
   * invariants are all satisfiable by seeing nothing, so without this one a
   * blind spot passes the suite: a change in a part the engine never reads
   * survives accept and reject alike, and self-compares clean.
   */
  "difference-is-reported",
] as const);

export type Invariant = (typeof INVARIANTS)[number];

export type InvariantOutcome =
  | { invariant: Invariant; status: "passed" }
  | { invariant: Invariant; status: "failed"; detail: string }
  | { invariant: Invariant; status: "skipped"; detail: string };

const sha256 = async (buffer: ArrayBuffer): Promise<string> => {
  const digest = await crypto.subtle.digest("SHA-256", buffer);
  return [...new Uint8Array(digest)].map((byte) => byte.toString(16).padStart(2, "0")).join("");
};

const sha256Text = async (text: string): Promise<string> =>
  await sha256(new TextEncoder().encode(text).buffer as ArrayBuffer);

export type CompareDigests = {
  /** The redlined package's bytes. */
  buffer: string;
  /** The change list, as canonical JSON. */
  changes: string;
};

export const digestsOf = async (
  buffer: ArrayBuffer,
  changes: readonly CompareChange[],
): Promise<CompareDigests> => ({
  buffer: await sha256(buffer),
  changes: await sha256Text(JSON.stringify(changes)),
});

type ResolveView = "accept" | "reject";

const resolvedBufferOf = async (redlined: ArrayBuffer, view: ResolveView): Promise<ArrayBuffer> => {
  const reviewer = await FolioDocxReviewer.fromBuffer(redlined);
  if (view === "accept") {
    reviewer.acceptAll();
  } else {
    reviewer.rejectAll();
  }
  return await reviewer.toBuffer();
};

type ChangeCount = { count: number } | { error: string };

const changeCountBetween = async (
  left: ArrayBuffer,
  right: ArrayBuffer,
  options: CompareDocxOptions,
): Promise<ChangeCount> => {
  const result = await compareDocx(left, right, options);
  return result.isErr()
    ? { error: `${result.error.name}: ${result.error.message}` }
    : { count: result.value.changes.length };
};

/**
 * The round-trip algebra, as a comparison rather than a text equality: the two
 * documents are the same document exactly when comparing them reports nothing.
 * Stating it that way also exercises the engine on its own output, which is
 * where a redline that reads plausibly and is wrong shows up.
 */
const checkRoundTrip = async (
  original: ArrayBuffer,
  redlined: ArrayBuffer,
  view: ResolveView,
  options: CompareDocxOptions,
): Promise<InvariantOutcome> => {
  const invariant: Invariant = view === "accept" ? "accept-returns-target" : "reject-returns-base";
  const resolved = await resolvedBufferOf(redlined, view);
  const outcome = await changeCountBetween(original, resolved, options);
  if ("error" in outcome) {
    return { invariant, status: "failed", detail: outcome.error };
  }
  return outcome.count === 0
    ? { invariant, status: "passed" }
    : {
        invariant,
        status: "failed",
        detail: `${String(outcome.count)} changes remain after ${view}All`,
      };
};

type DifferenceOptions = {
  expectation: "identical" | "different";
  changes: readonly CompareChange[];
  unsupported: readonly string[];
};

const differenceIsReported = ({
  expectation,
  changes,
  unsupported,
}: DifferenceOptions): InvariantOutcome => {
  const invariant: Invariant = "difference-is-reported";
  if (expectation === "identical") {
    return { invariant, status: "skipped", detail: "the pair is identical" };
  }
  if (changes.length > 0) {
    return { invariant, status: "passed" };
  }
  return {
    invariant,
    status: "failed",
    detail:
      unsupported.length === 0
        ? "no change and nothing reported unsupported"
        : `no change; unsupported: ${unsupported.join(", ")}`,
  };
};

const selfCompareIsEmpty = (outcome: ChangeCount): InvariantOutcome => {
  const invariant: Invariant = "self-compare-is-empty";
  if ("error" in outcome) {
    return { invariant, status: "failed", detail: outcome.error };
  }
  return outcome.count === 0
    ? { invariant, status: "passed" }
    : { invariant, status: "failed", detail: `${String(outcome.count)} changes against itself` };
};

const schemaValidity = (errors: readonly string[] | null): InvariantOutcome => {
  const invariant: Invariant = "schema-validity";
  if (errors === null) {
    return { invariant, status: "skipped", detail: "no Open XML SDK validator on this machine" };
  }
  return errors.length === 0
    ? { invariant, status: "passed" }
    : { invariant, status: "failed", detail: errors.slice(0, 5).join(" | ") };
};

export type CheckInvariantsOptions = {
  base: ArrayBuffer;
  target: ArrayBuffer;
  redlined: ArrayBuffer;
  changes: readonly CompareChange[];
  /** Reasons the comparison gave for parts it did not look at. */
  unsupported: readonly string[];
  /** Whether the harness built this pair to differ. */
  expectation: "identical" | "different";
  options: CompareDocxOptions;
  /** Validates a package against the OOXML schema, when a validator is available. */
  validate: PackageValidator | null;
};

export type InvariantReport = {
  outcomes: readonly InvariantOutcome[];
  digests: CompareDigests;
};

export const checkInvariants = async ({
  base,
  target,
  redlined,
  changes,
  unsupported,
  expectation,
  options,
  validate,
}: CheckInvariantsOptions): Promise<InvariantReport> => {
  const outcomes: InvariantOutcome[] = [
    await checkRoundTrip(base, redlined, "reject", options),
    await checkRoundTrip(target, redlined, "accept", options),
  ];

  outcomes.push(differenceIsReported({ expectation, changes, unsupported }));

  outcomes.push(selfCompareIsEmpty(await changeCountBetween(base, base, options)));

  const digests = await digestsOf(redlined, changes);
  const repeated = await compareDocx(base, target, options);
  if (repeated.isErr()) {
    outcomes.push({
      invariant: "byte-determinism",
      status: "failed",
      detail: `${repeated.error.name}: ${repeated.error.message}`,
    });
  } else {
    const again = await digestsOf(repeated.value.buffer, repeated.value.changes);
    outcomes.push(
      again.buffer === digests.buffer && again.changes === digests.changes
        ? { invariant: "byte-determinism", status: "passed" }
        : {
            invariant: "byte-determinism",
            status: "failed",
            detail: `buffer ${digests.buffer.slice(0, 12)} vs ${again.buffer.slice(0, 12)}, changes ${digests.changes.slice(0, 12)} vs ${again.changes.slice(0, 12)}`,
          },
    );
  }

  outcomes.push(schemaValidity(validate === null ? null : validate(redlined)));

  return { outcomes, digests };
};
