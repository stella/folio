import { describe, expect, test } from "bun:test";

import type { CensusSignature } from "./lib/corpus-census";
import {
  type ExpectedDispositions,
  compareToExpectedDispositions,
  dispositionOf,
  pathMatchesPattern,
  partitionExpectedDispositions,
  refreshedExpectedDispositions,
  validateExpectedDispositions,
  withoutDispositions,
} from "./lib/corpus-dispositions";
import { EXPECTED_DISPOSITIONS_PATH, EXPECTED_REFUSALS_PATH } from "./lib/corpus-manifest";
import type { ExpectedRefusals } from "./lib/corpus-refusals";
import { emptyCensus } from "./lib/corpus-census";
import { failureSignature } from "./lib/corpus-signature";

const EDITOR_ROUND_TRIP = "editor-round-trip";
const RESERIALIZE = "reserialize";

const signature = (invariant: string, message: string, files: number): CensusSignature => {
  const failure = { invariant, message, frame: "-" } as CensusSignature;
  return { ...failure, signature: failureSignature(failure), files, examples: [] };
};

const entry = (override: Record<string, unknown> = {}) => ({
  id: "an-id",
  match: {
    kind: "path",
    invariants: [EDITOR_ROUND_TRIP],
    paths: ["package.**.content[run].preservedAttributes"],
  },
  reason: "a reason",
  contract: "docs/container-contract.md#somewhere",
  removalCondition: "a carrier that does not exist yet",
  fileHits: 1,
  ...override,
});

const listOf = (...entries: ReturnType<typeof entry>[]): ExpectedDispositions =>
  ({ schemaVersion: 1, entries }) as unknown as ExpectedDispositions;

const NO_REFUSALS: ExpectedRefusals = { schemaVersion: 1, entries: [] };

const RUN_REMAINDER = "package.**.content[run].preservedAttributes";

describe("a path pattern claims an owner, not a spelling", () => {
  test("`**` stands for any run of segments, including none", () => {
    expect(
      pathMatchesPattern(
        RUN_REMAINDER,
        "package.document.content[paragraph].content[run].preservedAttributes",
      ),
    ).toBe(true);
    expect(
      pathMatchesPattern(
        RUN_REMAINDER,
        "package.document.sections[].content[table].rows[tableRow].cells[tableCell].content[paragraph].content[run].preservedAttributes",
      ),
    ).toBe(true);
  });

  /**
   * The paragraph keeps its remainder through the editor, so the row that
   * reports a paragraph losing one is a defect. While a segment named only the
   * field, the pattern could only approximate "under a run" as "one
   * `content[]` deeper than a paragraph", and a paragraph one level deeper
   * than that shape was claimed by accident.
   */
  test("a paragraph's remainder is not a run's, at any depth", () => {
    expect(
      pathMatchesPattern(RUN_REMAINDER, "package.document.content[paragraph].preservedAttributes"),
    ).toBe(false);
    expect(
      pathMatchesPattern(
        RUN_REMAINDER,
        "package.document.content[blockSdt].content[paragraph].preservedAttributes",
      ),
    ).toBe(false);
  });

  test("`[*]` claims every kind under a field, `[]` only the kinds the model leaves unnamed", () => {
    const any = "package.**.content[*].preservedAttributes";
    expect(pathMatchesPattern(any, "package.document.content[run].preservedAttributes")).toBe(true);
    expect(pathMatchesPattern(any, "package.document.content[].preservedAttributes")).toBe(true);
    expect(pathMatchesPattern(any, "package.document.rows[tableRow].preservedAttributes")).toBe(
      false,
    );
    const untyped = "package.**.content[].preservedAttributes";
    expect(pathMatchesPattern(untyped, "package.document.content[].preservedAttributes")).toBe(
      true,
    );
    expect(pathMatchesPattern(untyped, "package.document.content[run].preservedAttributes")).toBe(
      false,
    );
  });

  // A message is capped, so a long path arrives cut. The same decision must
  // not split into a claimed row and an unclaimed one just because one file's
  // path is deeper than another's.
  test("a path cut by the message cap still matches through its leaf", () => {
    expect(
      pathMatchesPattern(
        RUN_REMAINDER,
        "package.document.rows[tableRow].cells[tableCell].content[run].preservedAttribu…",
      ),
    ).toBe(true);
    expect(
      pathMatchesPattern(
        "package.**.content[*].preservedAttributes",
        "package.document.rows[tableRow].content[run].preservedAttribu…",
      ),
    ).toBe(true);
    // `[*]` stands for every kind, so a segment cut inside the kind still
    // agrees with it as far as either goes.
    expect(pathMatchesPattern("package.content[*]", "package.content[ru…")).toBe(true);
    expect(pathMatchesPattern("package.content[*]", "package.rows[ta…")).toBe(false);
  });

  test("a path cut before its leaf matches nothing", () => {
    expect(
      pathMatchesPattern(RUN_REMAINDER, "package.document.content[paragraph].cells[tableCe…"),
    ).toBe(false);
  });

  /**
   * A path too long for the cap keeps its head and its leaf and drops the
   * middle. The dropped run reads as `…`, which is what `**` already means, so
   * a decision anchored on its carrier keeps claiming its own rows however
   * deep the corpus buried them.
   */
  test("a path shortened from the middle still matches through `**`", () => {
    expect(pathMatchesPattern(RUN_REMAINDER, "package.….content[run].preservedAttributes")).toBe(
      true,
    );
    expect(
      pathMatchesPattern(
        "package.content[paragraph].content[run].preservedAttributes",
        "package.….content[run].preservedAttributes",
      ),
    ).toBe(false);
  });

  test("a pattern is restricted to the invariants whose leg the decision is about", () => {
    const dispositions = listOf(entry());
    const message =
      "editor round trip changed package.document.content[paragraph].content[run].preservedAttributes: array became absent";
    expect(dispositionOf(dispositions, signature(EDITOR_ROUND_TRIP, message, 1))).toBeDefined();
    expect(dispositionOf(dispositions, signature(RESERIALIZE, message, 1))).toBeUndefined();
  });
});

describe("the check reports dispositions rather than ratcheting them as defects", () => {
  const claimed = signature(
    EDITOR_ROUND_TRIP,
    "editor round trip changed package.document.content[paragraph].content[run].preservedAttributes: array became absent",
    7,
  );
  const defect = signature(
    EDITOR_ROUND_TRIP,
    "editor round trip changed package.document.content[paragraph].preservedAttributes: array became absent",
    2,
  );

  test("a claimed row leaves the census the baseline ratchets against", () => {
    const census = { ...emptyCensus("t1", "r1"), signatures: [claimed, defect] };
    const { defects, dispositions } = partitionExpectedDispositions(
      census,
      listOf(entry({ fileHits: 7 })),
    );
    expect(defects.signatures).toEqual([defect]);
    expect(dispositions).toEqual([claimed]);
  });

  // Both sides of every comparison drop the same rows, which is what lets an
  // entry land before the baselines are re-measured.
  test("the same rows leave a committed baseline", () => {
    expect(withoutDispositions([claimed, defect], listOf(entry({ fileHits: 7 })))).toEqual([
      defect,
    ]);
  });

  test("growth fails: the class widened past the decision", () => {
    expect(compareToExpectedDispositions(listOf(entry({ fileHits: 6 })), [claimed])).toEqual([
      {
        kind: "more-files",
        signature: "an-id",
        detail:
          "7 file hits carry this disposition, the entry allows 6; the class widened past docs/container-contract.md#somewhere",
      },
    ]);
  });

  test("a shrink is written down rather than absorbed", () => {
    expect(
      compareToExpectedDispositions(listOf(entry({ fileHits: 9 })), [claimed]).map(
        ({ kind }) => kind,
      ),
    ).toEqual(["fewer-files"]);
  });

  test("an entry nothing matches any more has to be deleted", () => {
    expect(
      compareToExpectedDispositions(listOf(entry({ fileHits: 7 })), [defect]).map(
        ({ kind }) => kind,
      ),
    ).toEqual(["resolved-disposition"]);
  });

  test("a refresh moves the counts and nothing else", () => {
    expect(refreshedExpectedDispositions(listOf(entry({ fileHits: 1 })), [claimed])).toEqual(
      listOf(entry({ fileHits: 7 })),
    );
  });
});

describe("the committed list is a decision a reviewer can weigh", () => {
  test("every entry states a reason, a contract and a removal condition", async () => {
    const [dispositions, refusals] = await Promise.all([
      Bun.file(EXPECTED_DISPOSITIONS_PATH).json(),
      Bun.file(EXPECTED_REFUSALS_PATH).json() as Promise<ExpectedRefusals>,
    ]);
    expect(validateExpectedDispositions(dispositions, refusals)).toEqual([]);
    expect(dispositions.entries.length).toBeGreaterThan(0);
  });

  test("a missing hand-written field is an issue, field by field", () => {
    for (const field of ["reason", "contract", "removalCondition"]) {
      expect(validateExpectedDispositions(listOf(entry({ [field]: "  " })), NO_REFUSALS)).toEqual([
        `entries[0].${field}: expected a non-empty ${field}`,
      ]);
    }
  });

  // The two lists answer different questions, and a signature in both would be
  // excused twice and ratcheted under two rules at once.
  test("an entry that also claims an expected refusal is refused", () => {
    const refusalMessage =
      "editor round trip changed package.document.content[paragraph].content[run].preservedAttributes: array became absent";
    const refusals: ExpectedRefusals = {
      schemaVersion: 1,
      entries: [
        {
          signature: failureSignature({
            invariant: EDITOR_ROUND_TRIP,
            message: refusalMessage,
            frame: "-",
          } as Parameters<typeof failureSignature>[0]),
          reason: "a refusal",
          files: 1,
        },
      ],
    };
    expect(validateExpectedDispositions(listOf(entry()), refusals)).toEqual([
      `an-id also claims the expected refusal ${refusals.entries[0]?.signature}; a signature is a refusal or a disposition, never both`,
    ]);
  });

  /**
   * A mistyped kind claims nothing, and a disposition that claims nothing
   * surfaces a nightly later as an entry to delete rather than as a typo. The
   * vocabulary is closed and committed, so it is checked where the file is
   * read.
   */
  test("a kind the model does not declare is an issue, not a silent miss", () => {
    const withKind = (kind: string) =>
      listOf(
        entry({
          match: {
            kind: "path",
            invariants: [EDITOR_ROUND_TRIP],
            paths: [`package.**.content[${kind}].preservedAttributes`],
          },
        }),
      );
    expect(validateExpectedDispositions(withKind("paragrpah"), NO_REFUSALS)).toEqual([
      "entries[0].match.paths[0]: `paragrpah` is not a kind the model declares; a pattern that names none claims nothing",
    ]);
    for (const kind of ["run", "*", ""]) {
      expect(validateExpectedDispositions(withKind(kind), NO_REFUSALS)).toEqual([]);
    }
  });

  test("two entries may not share an id", () => {
    expect(validateExpectedDispositions(listOf(entry(), entry()), NO_REFUSALS)).toEqual([
      "entries[1].id: duplicate id `an-id`",
    ]);
  });
});
