/**
 * `explain` says what it found, including when it found nothing to measure.
 *
 * A pair the law could not run and a pair that survived both carry
 * `mechanism: null`, so the mechanism alone printed "survives" for every pair
 * in a part a repack copies through — `word/styles.xml` among them. That is the
 * first line a reader of the class hits, and it said the opposite of the truth.
 */

import { describe, expect, test } from "bun:test";

import { allSubjects, explainedOutcome } from "./container-survival-census";
import type { Subject } from "./lib/container-survival/fixture";
import { runSurvivalLaws, subjectKey } from "./lib/container-survival/laws";
import { loadContainerSpace, WML_NAMESPACE } from "./lib/container-survival/schemaSpace";

const space = await loadContainerSpace();

const subjectAt = (key: string): Subject => {
  const found = allSubjects(space).find(
    (subject) => subjectKey(subject).replaceAll(`{${WML_NAMESPACE}}`, "w:") === key,
  );
  if (found === undefined) {
    throw new Error(`no pair in the census space is keyed ${key}`);
  }
  return found;
};

const outcomeLineOf = async (key: string): Promise<string> =>
  explainedOutcome(await runSurvivalLaws(space, subjectAt(key)));

describe("the outcome line separates unmeasured from unharmed", () => {
  test("a pair in a part a repack copies through prints why it was not measured", async () => {
    expect(await outcomeLineOf("w:pPr|w:CT_PPrGeneral/w:cnfStyle")).toBe(
      "unrepresentable: a repack replays word/styles.xml verbatim, " +
        "and folio splices style definitions into it rather than rebuilding it",
    );
  }, 60_000);

  test("a pair in a part the law rebuilds is measured rather than skipped", async () => {
    expect(await outcomeLineOf("w:font|w:CT_Font/w:notTrueType")).toBe("mechanism: survives");
  }, 60_000);

  test("a pair that came back still prints that it survives", async () => {
    expect(await outcomeLineOf("w:tblBorders|w:CT_TblBorders/w:top")).toBe("mechanism: survives");
  }, 60_000);

  test("a pair that was lost still prints its mechanism", async () => {
    expect(await outcomeLineOf("w:p|w:CT_P/w:pPr")).toBe(
      "mechanism: serialized-only-via-verbatim-replay",
    );
  }, 60_000);
});
