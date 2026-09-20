/**
 * The one disposition nothing runs.
 *
 * A handler and a `CAPTURE` are verified by execution: the survival law builds
 * a fixture, saves it, and reports what came back. `OWNED_ELSEWHERE` is not a
 * behaviour, it is a **claim about another module** — "some other reader takes
 * this child and writes it back, so this walk must not capture it as well" —
 * and nothing checked that the other reader existed. `w:tr/w:tblPrEx` carried
 * the claim while the contract recorded the same pair as `dropped`
 * (`neverParsed`), the two statements sat beside each other, and neither side
 * knew the other was there.
 *
 * So the claim is data. `ownedElsewhere` in `docx/containerChildren.ts` is the
 * only way to make one, it names the reader as `<module>#<export>`, and it
 * registers what it made. This module reads that registry and asks two
 * questions of it:
 *
 * - **Does the owner exist?** The module is imported and the export looked up,
 *   so a reader that was renamed, moved or never written fails the check
 *   rather than sitting in a comment.
 * - **Does the contract agree?** A pair an owner claims may not be recorded
 *   `dropped (neverParsed)` or `dropped (containerNotKept)`: the first says no
 *   parser reads the markup, which is the direct contradiction, and the second
 *   says the pair went with a container folio does not keep, in which case the
 *   owner cannot have written it back either.
 *
 * The other drop reasons are not contradictions. `replayOnly`,
 * `editorProjection` and their neighbours all describe markup a reader *did*
 * take and a later stage lost, which is what an owner claim says happened.
 *
 * A claim is registered when the module that makes it loads, so the check
 * loads them first — by scanning the sources for the call, not by keeping a
 * list of them beside this file.
 */

import path from "node:path";

import { Glob } from "bun";

import {
  type ChildOwnerClaim,
  ownedElsewhereClaims,
} from "@stll/folio-core/docx/containerChildren";
import { TaggedError } from "better-result";

import {
  type ContainerContract,
  DISPOSITIONS,
  type DropReason,
} from "../../../specifications/container-contract/dispositions";
import { DISPATCHED_CONTAINERS } from "./dispatchedContainers";
import { containerKey, qualify, WML_NAMESPACE } from "./schemaSpace";

const REPOSITORY_ROOT = path.resolve(import.meta.dir, "../../..");
const PACKAGES_ROOT = path.join(REPOSITORY_ROOT, "packages");
const CORE_DOCX = path.join(PACKAGES_ROOT, "core/src/docx");

class ContainerOwnershipError extends TaggedError("ContainerOwnershipError")<{
  message: string;
}> {}

/**
 * The drop reasons an owner claim contradicts.
 *
 * Total over `DropReason` so a new reason class has to be classified here
 * rather than defaulting to "not a contradiction", which is the direction that
 * loses a finding.
 */
const CONTRADICTS_AN_OWNER = {
  containerNotKept: true,
  neverParsed: true,
  parsedNotSerialized: false,
  replayOnly: false,
  replayRejected: false,
  editorProjection: false,
  respelled: false,
  parserThrows: false,
} as const satisfies Record<DropReason, boolean>;

const wml = (name: string): string => qualify({ namespace: WML_NAMESPACE, name });

const pairKey = (element: string, type: string, child: string): string =>
  `${containerKey({
    element: { namespace: WML_NAMESPACE, name: element },
    typeQName: wml(type),
  })}/${wml(child)}`;

/**
 * Pairs an owner claims that the contract records as lost anyway, and why.
 *
 * A ratchet, not an exemption: the check fails on a pair that is not here, and
 * fails on a pair here that no longer violates, so the list can only shrink.
 * Each entry names the mechanism, because "the owner reads it" and "the census
 * loses it" are both true and what sits between them is the finding.
 */
const KNOWN_OWNED_LOSSES: Readonly<Record<string, string>> = {
  [pairKey("tr", "CT_Row", "trPr")]:
    "The owner is real and a stated row property survives. `parseTableRowProperties` returns " +
    "`undefined` when nothing in the element was modelled, so a `w:trPr` that states nothing — " +
    "which is what the census's fixture builds — reaches no model and no serializer writes one " +
    "back. Closing it means giving the element a carrier of its own instead of one keyed on the " +
    "properties it yielded.",
  [pairKey("tc", "CT_Tc", "tcPr")]:
    "The same shape one level down: `parseTableCellProperties` returns `undefined` for a `w:tcPr` " +
    "that yields no typed property, so the empty fixture is lost while a stated cell property " +
    "survives.",
};

const claimedContainers = (claim: ChildOwnerClaim): string[] => {
  const row = DISPATCHED_CONTAINERS.find(({ key }) => key === claim.container);
  if (row === undefined) {
    throw new ContainerOwnershipError({
      message: `no dispatched container named ${claim.container}`,
    });
  }
  return row.members.map(([element, type]) =>
    containerKey({
      element: { namespace: WML_NAMESPACE, name: element },
      typeQName: qualify({ namespace: WML_NAMESPACE, name: type }),
    }),
  );
};

/**
 * The census pair keys one claim covers.
 *
 * A handler map serves every member of its dispatcher row — one walk reads a
 * `w:tr` and the `w:sdtContent` of a row-level content control — so a claim
 * made in that map is made about each of them. A member whose content model
 * does not declare the child simply has no pair, and the contract lookup
 * misses it.
 */
export const claimedPairs = (claim: ChildOwnerClaim): string[] =>
  claimedContainers(claim).map((container) => `${container}/${wml(claim.child)}`);

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null;

/** Every file under `packages/*​/src` that constructs an owner claim. */
const claimingModules = async (): Promise<string[]> => {
  const files: string[] = [];
  for await (const file of new Glob("*/src/**/*.ts").scan({
    absolute: true,
    cwd: PACKAGES_ROOT,
  })) {
    if (file.endsWith(".test.ts")) {
      continue;
    }
    if ((await Bun.file(file).text()).includes("ownedElsewhere(")) {
      files.push(file);
    }
  }
  return files.toSorted();
};

/**
 * Load the claiming modules, then hand back what they registered.
 *
 * Reading the registry without this returns whatever happened to be imported,
 * and a check that silently sees no claims passes for the wrong reason.
 */
export const loadOwnerClaims = async (): Promise<readonly ChildOwnerClaim[]> => {
  for (const file of await claimingModules()) {
    await import(file);
  }
  const claims = ownedElsewhereClaims();
  if (claims.length === 0) {
    throw new ContainerOwnershipError({
      message: "no owner claims were registered; the scan found no module that makes one",
    });
  }
  return claims;
};

/** Resolve `<module>#<export>` against `packages/core/src/docx`. */
const unresolved = async (reader: string): Promise<string | undefined> => {
  const [module, exported] = reader.split("#");
  if (module === undefined || exported === undefined || module === "" || exported === "") {
    return `${reader} is not a <module>#<export> reference`;
  }
  const file = path.join(CORE_DOCX, `${module}.ts`);
  if (!(await Bun.file(file).exists())) {
    return `${reader} names no module: ${path.relative(REPOSITORY_ROOT, file)} does not exist`;
  }
  const namespace: unknown = await import(file);
  if (!isRecord(namespace) || typeof namespace[exported] !== "function") {
    return `${reader} names no exported function ${exported}`;
  }
  return undefined;
};

/**
 * The claims whose reader is not there, with what was looked for.
 *
 * The module is imported rather than read, so a reference that survives a
 * rename only because a comment was not updated fails here.
 */
export const unresolvedOwners = async (claims: readonly ChildOwnerClaim[]): Promise<string[]> => {
  const failures: string[] = [];
  for (const claim of claims) {
    const failure = await unresolved(claim.reader);
    if (failure !== undefined) {
      failures.push(
        `${claim.container}/${claim.child} claims an owner that does not resolve: ${failure}`,
      );
    }
  }
  return failures;
};

/**
 * Every disagreement between the owner claims and the committed contract.
 *
 * Reads the contract rather than running the census: the claims and the
 * recorded dispositions are both data, so the check costs nothing and runs
 * under a scoped census too.
 */
export const ownershipViolations = async (contract: ContainerContract): Promise<string[]> => {
  const claims = await loadOwnerClaims();
  const problems = await unresolvedOwners(claims);
  const violating = new Set<string>();

  for (const claim of claims) {
    for (const key of claimedPairs(claim)) {
      const entry = contract.entries[key];
      if (entry === undefined || entry.disposition !== DISPOSITIONS.dropped) {
        continue;
      }
      if (!CONTRADICTS_AN_OWNER[entry.reason]) {
        continue;
      }
      violating.add(key);
      if (KNOWN_OWNED_LOSSES[key] === undefined) {
        problems.push(
          `${claim.reader} claims to own it, the contract drops it (${entry.reason}): ${key}`,
        );
      }
    }
  }

  for (const key of Object.keys(KNOWN_OWNED_LOSSES)) {
    if (!violating.has(key)) {
      problems.push(
        `a known owned loss no longer happens; drop it from KNOWN_OWNED_LOSSES: ${key}`,
      );
    }
  }

  return problems;
};
