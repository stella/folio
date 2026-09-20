/**
 * `OWNED_ELSEWHERE` used to be a word, and now it is a claim with an address.
 *
 * A handler and a `CAPTURE` are checked by the survival law running them. An
 * owner claim is a statement about another module, and for `w:tr/w:tblPrEx`
 * there was no other module: the child walk skipped the element, nothing read
 * it, and the container contract recorded the same pair as never parsed
 * without either side noticing the other.
 *
 * These three tests are the other side. The first imports every reader a claim
 * names, so a rename cannot leave the claim pointing at nothing. The second
 * holds the claims against the committed contract. The third refuses a claim
 * that has stopped being about any pair at all, which is how a claim survives
 * the element it was written for.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import path from "node:path";

import type { ContainerContract } from "../specifications/container-contract/dispositions";
import {
  claimedPairs,
  loadOwnerClaims,
  ownershipViolations,
  unresolvedOwners,
} from "./lib/container-survival/ownership";

setDefaultTimeout(60_000);

const REPOSITORY_ROOT = path.resolve(import.meta.dir, "..");

const contract = async (): Promise<ContainerContract> =>
  Bun.file(path.join(REPOSITORY_ROOT, "specifications/container-contract/contract.json")).json();

describe("container child ownership", () => {
  test("every reader a handler map claims is an exported function", async () => {
    expect(await unresolvedOwners(await loadOwnerClaims())).toEqual([]);
  });

  test("the contract records no owned pair as never parsed or lost with its container", async () => {
    expect(await ownershipViolations(await contract())).toEqual([]);
  });

  test("every claim covers at least one pair the contract decides", async () => {
    const { entries } = await contract();
    const empty = (await loadOwnerClaims())
      .filter((claim) => !claimedPairs(claim).some((key) => entries[key] !== undefined))
      .map((claim) => `${claim.container}/${claim.child}`);
    expect(empty).toEqual([]);
  });
});
