import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import { assertProperty, propertyTestTimeout } from "../test/property-testing";
import {
  failureMarker,
  failureRecord,
} from "../test/consumer-scenarios/support/failure-fingerprints";
import { failureClass, normalizeFailureMessage } from "./fuzz-failure-class";
import type { Context, Finding } from "./fuzz-failure-issues";
import {
  type ClosedIssue,
  type ClosedIssueDisposition,
  duplicateTarget,
  fileClasses,
  type Issue,
  type IssueStore,
  MAX_NEW_ISSUES,
  NewIssueCapError,
} from "./fuzz-issue-classes";
import flood from "./fuzz-report-flood.fixtures.json";

const context = {
  runUrl: "https://github.com/stella/folio/actions/runs/42",
  sha: "abc123",
  source: "continuous fuzz",
  date: "2026-10-02T12:00:00Z",
} satisfies Context;

const finding = (seed: number, message = "listLevel is 3, expected 2"): Finding => {
  const error = new Error(message);
  const marker = failureMarker({
    test: "consumer flow lists / direct",
    seed,
    path: `0:${seed}`,
    repro: `PROPERTY_TEST_SEED=${seed} bun scripts/consumer-scenarios.ts`,
    failure: error,
  });
  const record = failureRecord(marker, error);
  return { record, seeds: [seed], records: [record] };
};

const fakeStore = (initial: Issue[] = []) => {
  const issues = initial.map((issue) => ({ ...issue }));
  const edits: number[] = [];
  const reopens: number[] = [];
  const comments: { number: number; comment: string }[] = [];
  const created: number[] = [];
  const requireIssue = (number: number) => {
    const issue = issues.find((candidate) => candidate.number === number);
    if (issue === undefined) throw new Error("Missing fake issue");
    return issue;
  };
  const store = {
    list: async () => issues.map((issue) => ({ ...issue })),
    create: async (title: string, body: string) => {
      const number = Math.max(0, ...issues.map((issue) => issue.number)) + 1;
      const issue = { number, title, body, state: "open", closedAt: null } satisfies Issue;
      issues.push(issue);
      created.push(number);
      return { ...issue };
    },
    edit: async (number: number, body: string) => {
      requireIssue(number).body = body;
      edits.push(number);
    },
    reopen: async (number: number, comment: string) => {
      const issue = requireIssue(number);
      const { title, body } = issue;
      issues[issues.indexOf(issue)] = { number, title, body, state: "open", closedAt: null };
      reopens.push(number);
      comments.push({ number, comment });
    },
  } satisfies IssueStore;
  return { store, issues, edits, reopens, comments, created };
};

const FIXED = { type: "fixed" } as const satisfies ClosedIssueDisposition;
const SUPERSEDED = { type: "superseded" } as const satisfies ClosedIssueDisposition;

const report = (fake: ReturnType<typeof fakeStore>, findings: Finding[], runContext = context) =>
  fileClasses({ findings, context: runContext, store: fake.store, known: new Map() });

const legacy = (closedAt: string | null = null): Issue => {
  const fields = {
    number: 20,
    title:
      "Fuzz failure [0123456789abcdef]: consumer flow lists / direct: listLevel is 3, expected 2",
    body: "Legacy failure details",
  };
  return closedAt === null
    ? { ...fields, state: "open", closedAt }
    : { ...fields, state: "closed", closedAt, disposition: FIXED };
};

describe("class issue filing", () => {
  test("same class retains every seed/path/replay and counts distinct runs", async () => {
    const fake = fakeStore();
    const first = finding(1);
    const second = finding(2, "listLevel is undefined, expected 4");
    second.record.marker = failureMarker({
      ...second.record.marker,
      flow: "insertText → save",
      failure: new Error("listLevel is 7, expected 4"),
    });
    expect(second.record.marker.fingerprint).not.toBe(first.record.marker.fingerprint);
    second.record.replays.push("PROPERTY_TEST_SEED=2 bun scripts/replay.ts");
    await report(fake, [first, second]);
    expect(fake.created).toEqual([1]);
    const body = fake.issues.at(0)?.body;
    expect(body).toContain("Seen in 1 distinct runs; 2 replay rows.");
    expect(body).toContain(first.record.marker.fingerprint);
    expect(body).toContain(second.record.marker.fingerprint);
    expect(body).toContain("0:1");
    expect(body).toContain("0:2");
    expect(body).toContain("PROPERTY_TEST_SEED=2 bun scripts/replay.ts");
    await report(fake, [first, second]);
    expect(fake.edits).toEqual([]);
    await report(fake, [finding(3)]);
    expect(fake.edits).toEqual([1]);
    expect(fake.issues.at(0)?.body).toContain("Seen in 1 distinct runs");
    await report(fake, [first], {
      ...context,
      runUrl: "https://github.com/stella/folio/actions/runs/43",
    });
    expect(fake.issues.at(0)?.body).toContain("Seen in 2 distinct runs");
    expect(fake.created).toEqual([1]);
  });

  test("different classes create separate issues", async () => {
    const fake = fakeStore();
    await report(fake, [finding(1)]);
    const result = await report(fake, [finding(2), finding(3, "unrelated failure")]);
    expect(fake.created).toEqual([1, 2]);
    expect(fake.edits).toEqual([1]);
    expect(result.map(({ issue }) => issue)).toEqual(["#1", "#2"]);
    expect(fake.issues.map(({ title }) => title)).toEqual([
      "List numbering: listLevel mismatch",
      "Consumer flow: unrelated failure",
    ]);
  });

  test("recent closed class reopens with regression note, then rerun changes nothing", async () => {
    const fake = fakeStore([legacy("2026-09-18T12:00:01Z")]);
    await report(fake, [finding(1)]);
    expect(fake.reopens).toEqual([20]);
    expect(fake.comments.at(0)?.comment).toBe(
      [
        `Reopening: this class recurred in ${context.runUrl}.`,
        "",
        `Recurring fingerprints: \`${finding(1).record.marker.fingerprint}\``,
      ].join("\n"),
    );
    expect(fake.created).toEqual([]);
    expect(fake.issues.at(0)?.body).toContain("Regressed:");
    expect(fake.issues.at(0)?.body).toContain("Legacy failure details");
    await report(fake, [finding(1)]);
    expect(fake.edits).toEqual([20]);
    expect(fake.reopens).toEqual([20]);
  });

  for (const closedAt of ["2026-09-18T12:00:00Z", "2026-09-17T12:00:00Z"]) {
    test(`closed class at least 14 days old creates linked issue (${closedAt})`, async () => {
      const fake = fakeStore([legacy(closedAt)]);
      await report(fake, [finding(1)]);
      expect(fake.reopens).toEqual([]);
      expect(fake.created).toEqual([21]);
      expect(fake.issues.at(1)?.body).toContain("Previous occurrence: #20");
    });
  }

  test("any open legacy issue beats a closed marker issue; lowest open wins", async () => {
    const signature = failureClass("consumer flow lists / direct", "listLevel is 3, expected 2");
    const marker = `<!-- fuzz-class: ${signature.key} -->`;
    const fake = fakeStore([
      { ...legacy(context.date), number: 19, body: marker },
      { ...legacy(), number: 21 },
      legacy(),
      { ...legacy(), number: 22, body: marker },
    ]);
    const result = await report(fake, [finding(1)]);
    expect(result.map(({ issue }) => issue)).toEqual(["#20"]);
    expect(fake.edits).toEqual([20]);
    expect(fake.reopens).toEqual([]);
    expect(fake.created).toEqual([]);
  });

  test("latest closed occurrence is selected by closure time", async () => {
    const fake = fakeStore([
      { ...legacy("2026-10-01T12:00:00Z"), number: 19 },
      legacy("2026-09-01T12:00:00Z"),
    ]);
    const result = await report(fake, [finding(1)]);
    expect(result.at(0)?.issue).toBe("#19");
    expect(fake.reopens).toEqual([19]);
  });

  test("known fingerprint and primary registry entries do not create issues", async () => {
    const fake = fakeStore();
    const known = finding(1);
    const result = await fileClasses({
      findings: [known],
      context,
      store: fake.store,
      known: new Map([[known.record.marker.fingerprint, "#100"]]),
    });
    expect(result.at(0)?.issue).toBe("#100");
    expect(fake.created).toEqual([]);
    known.record.marker.primary = "0123456789abcdef";
    const primary = await fileClasses({
      findings: [known],
      context,
      store: fake.store,
      known: new Map([["0123456789abcdef", "#101"]]),
    });
    expect(primary.at(0)?.issue).toBe("#101");
  });

  test("human legacy finding matches exact fingerprint without collapsing unrelated generic failures", async () => {
    const first = finding(1, "expect(received).toEqual(expected)");
    const human = {
      ...legacy(),
      title: "Revision acceptance: accepted tracked batches differ from direct batches",
      body: `Historical trace\n<!-- fuzz-failure-state ${JSON.stringify({ fingerprint: first.record.marker.fingerprint, primary: "0123456789abcdef" })} -->`,
    };
    const fake = fakeStore([human]);
    const result = await report(fake, [
      first,
      finding(2, "expect(received).toEqual(expected) with something different"),
    ]);
    expect(result.at(0)?.issue).toBe("#20");
    expect(result.at(1)?.issue).toBe("#21");
    expect(fake.issues.at(0)?.body).toContain("Historical trace");
    expect(fake.issues.at(0)?.body).toContain("| lists |  | direct |");
    await report(fake, [first]);
    expect(fake.edits).toEqual([20]);
  });

  test("one record known to a legacy issue takes the whole class there", async () => {
    for (const reverse of [false, true]) {
      const first = finding(1, "expect(received).toEqual(expected)");
      const second = finding(2, "expect(received).toEqual(expected)");
      expect(second.record.marker.fingerprint).toBe(first.record.marker.fingerprint);
      first.record.marker.primary = "0123456789abcdef";
      const human = {
        ...legacy(),
        title: "Revision acceptance: accepted tracked batches differ from direct batches",
        body: '<!-- fuzz-failure-state {"fingerprint":"0123456789abcdef"} -->',
      };
      const fake = fakeStore([human]);
      const records = [first.record, second.record];
      const combined = {
        record: first.record,
        seeds: [1, 2],
        records: reverse ? records.toReversed() : records,
      } satisfies Finding;
      const result = await report(fake, [combined]);
      expect(result).toEqual([{ fingerprint: first.record.marker.fingerprint, issue: "#20" }]);
      expect(fake.created).toEqual([]);
      expect(fake.edits).toEqual([20]);
      expect(fake.issues.at(0)?.body).toContain("2 replay rows");
    }
  });

  test("a legacy fingerprint never pulls an issue away from the class it is marked with", async () => {
    const item = finding(1, "unrelated failure");
    const other = failureClass("consumer flow lists / direct", "listLevel is 3, expected 2");
    const marked = {
      ...legacy(),
      body: `<!-- fuzz-class: ${other.key} -->\n<!-- fuzz-failure-state ${JSON.stringify({ fingerprint: item.record.marker.fingerprint })} -->`,
    };
    const fake = fakeStore([marked]);
    await report(fake, [item]);
    expect(fake.edits).toEqual([]);
    expect(fake.created).toEqual([21]);
  });

  test("replay text cannot terminate hidden state and split an idempotent class", async () => {
    const fake = fakeStore();
    const item = finding(1);
    item.record.replays = ["bun replay.ts --text '<!-- hidden -->'"];
    item.record.error = "literal --> in assertion";
    await report(fake, [item]);
    await report(fake, [item]);
    expect(fake.created).toEqual([1]);
    expect(fake.edits).toEqual([]);
    expect(fake.issues.at(0)?.body).toContain("\\u003e");
  });

  test("malformed structured state preserves original evidence while rebuilding state", async () => {
    const signature = failureClass("consumer flow lists / direct", "listLevel is 3, expected 2");
    const evidence = `<!-- fuzz-class: ${signature.key} -->\n<!-- fuzz-class-state {"key":${JSON.stringify(signature.key)},"runs":[],"seeds":[{"seed":"bad"}]} -->`;
    const fake = fakeStore([{ ...legacy(), body: evidence }]);
    await report(fake, [finding(1)]);
    expect(fake.edits).toEqual([20]);
    expect(fake.issues.at(0)?.body).toContain(evidence);
  });
});

const failing = (failure: { test: string; message: string }, seed: number): Finding => {
  const error = new Error(failure.message);
  const marker = failureMarker({
    test: failure.test,
    seed,
    repro: `FOLIO_SCENARIO_SEED=${seed} bun scripts/consumer-scenarios.ts`,
    failure: error,
  });
  const record = failureRecord(marker, error);
  return { record, seeds: [seed], records: [record] };
};

const floodFindings = (): Finding[] =>
  flood.failures.map((failure, index) => failing(failure, 1_000 + index));

const escaped = (json: string): string =>
  json.replaceAll("<", "\\u003c").replaceAll(">", "\\u003e");

/** The class marker as it was written while the key held the fixture and the mode. */
const oldMarker = (signature: readonly string[]): string =>
  `<!-- fuzz-class: ${escaped(JSON.stringify(signature))} -->`;

const openIssue = (number: number, title: string, body = "Legacy failure details"): Issue => ({
  number,
  title,
  body,
  state: "open",
  closedAt: null,
});

const standingIssues = (): Issue[] =>
  flood.standing.map(({ number, title }) => openIssue(number, title));

const writes = (fake: ReturnType<typeof fakeStore>): number =>
  fake.created.length + fake.edits.length + fake.reopens.length;

const bodyOf = (fake: ReturnType<typeof fakeStore>, number: number): string =>
  fake.issues.find((issue) => issue.number === number)?.body ?? "";

const rowCount = (body: string): number =>
  body.split("\n").filter((line) => /^\| [0-9a-f]{16} \|/u.test(line)).length;

const snapshot = (fake: ReturnType<typeof fakeStore>): string => JSON.stringify(fake.issues);

const NUMBERING = "Consumer flow: Numbering provenance missing for <id>";
const REPACK = "Consumer flow: Cannot repack invalid DOCX document model";
const EQUALITY = "expect(received).toEqual(expected)";

describe("a failure seen across fixtures and modes", () => {
  test("the recorded run is eighteen findings that were filed as eighteen classes", () => {
    const findings = floodFindings();
    expect(findings).toHaveLength(18);
    expect(new Set(findings.map(({ record }) => record.marker.fingerprint)).size).toBe(18);
    // What each finding was keyed by when the key held the fixture and the mode.
    const filedAs = findings.map(({ record: { marker } }) => {
      const signature = failureClass(marker.test, marker.assertion);
      return JSON.stringify(
        signature.fixture === ""
          ? [normalizeFailureMessage(marker.test), marker.test, "all", signature.message]
          : ["consumer-flow", signature.fixture, signature.mode, signature.message],
      );
    });
    expect(new Set(filedAs)).toEqual(
      new Set(flood.filed.map(({ signature }) => JSON.stringify(signature))),
    );
    expect(new Set(filedAs).size).toBe(18);
  });

  test("files one issue per failure and adds rows to the open issues it already has", async () => {
    const fake = fakeStore(standingIssues());
    const result = await report(fake, floodFindings());

    const opened = fake.issues.filter(({ number }) => fake.created.includes(number));
    expect(opened.map(({ title }) => title).toSorted()).toEqual([
      REPACK,
      NUMBERING,
      `Fuzz checks: accepting every revision matches the direct edit: ${EQUALITY}`,
      `Fuzz checks: numbering survives save and reopen: ${EQUALITY}`,
      `Fuzz checks: round-trip preserves document structure: ${EQUALITY}`,
    ]);
    expect(fake.issues).toHaveLength(8);
    expect(fake.edits.toSorted()).toEqual([1384, 1389, 1409]);
    expect(fake.reopens).toEqual([]);

    // The fixtures and modes of a class are the rows of its one issue.
    expect(rowCount(bodyOf(fake, 1384))).toBe(6);
    expect(rowCount(bodyOf(fake, 1389))).toBe(2);
    expect(rowCount(bodyOf(fake, 1409))).toBe(1);
    for (const number of [1384, 1389, 1409]) {
      expect(bodyOf(fake, number)).toContain("Legacy failure details");
    }
    expect(bodyOf(fake, 1389)).toContain("| lists |  | suggested |");
    expect(bodyOf(fake, 1389)).toContain("| styleNumbered |  | direct |");
    const numbering = opened.find(({ title }) => title === NUMBERING)?.body ?? "";
    expect(rowCount(numbering)).toBe(5);
    for (const cells of [
      "| comments |  | tracked-changes |",
      "| plain |  | tracked-changes |",
      "| emoji |  | tracked-changes |",
      "| stories |  | tracked-changes |",
      "| tables |  | suggested |",
    ]) {
      expect(numbering).toContain(cells);
    }
    for (const { body } of opened.filter(({ title }) => title !== NUMBERING)) {
      expect(rowCount(body ?? "")).toBe(1);
    }

    // Every finding went to exactly one issue.
    expect(result).toHaveLength(18);
    expect(new Set(result.map(({ fingerprint }) => fingerprint)).size).toBe(18);
    expect(new Set(result.map(({ issue }) => issue)).size).toBe(8);

    const after = snapshot(fake);
    const written = writes(fake);
    await report(fake, floodFindings());
    expect(writes(fake)).toBe(written);
    expect(snapshot(fake)).toBe(after);
  });

  test("open issues that are now one class are left open; only the lowest takes rows", async () => {
    const earlierRun = "https://github.com/stella/folio/actions/runs/41";
    const earlierState = (signature: readonly string[]): string =>
      `<!-- fuzz-class-state ${escaped(
        JSON.stringify({
          key: JSON.stringify(signature),
          runs: [earlierRun],
          seeds: [
            {
              fingerprint: "a1904c0b553b8545",
              flow: "deleteBlock > setBlockParagraphProperties",
              mode: "tracked-changes",
              seed: 1_026_398_003,
              path: null,
              repro: "replay",
              replays: ["replay"],
              evidence: "collisions flow failed at step 14",
              firstSeenRun: earlierRun,
            },
          ],
        }),
      )} -->`;
    const fake = fakeStore([
      ...standingIssues(),
      ...flood.filed.map(({ number, title, signature }) =>
        openIssue(
          number,
          title,
          number === 1453
            ? `${oldMarker(signature)}\n${earlierState(signature)}`
            : oldMarker(signature),
        ),
      ),
    ]);
    const before = new Map(fake.issues.map(({ number, body }) => [number, body]));
    const result = await report(fake, floodFindings());

    expect(fake.created).toEqual([]);
    expect(fake.reopens).toEqual([]);
    const canonical = [1384, 1389, 1409, 1448, 1449, 1450, 1453, 1466];
    expect(fake.edits.toSorted()).toEqual(canonical);
    expect(fake.issues).toHaveLength(21);
    for (const issue of fake.issues) {
      expect(issue.state).toBe("open");
      if (!canonical.includes(issue.number)) {
        expect(before.has(issue.number)).toBe(true);
        expect(issue.body).toBe(before.get(issue.number) ?? null);
      }
    }
    expect(new Set(result.map(({ issue }) => issue))).toEqual(
      new Set(canonical.map((number) => `#${number}`)),
    );
    expect(rowCount(bodyOf(fake, 1384))).toBe(6);
    // Rows an issue already had under the earlier key are kept, not set aside.
    expect(bodyOf(fake, 1453)).toContain("Seen in 2 distinct runs; 6 replay rows.");
    expect(bodyOf(fake, 1453)).toContain("| a1904c0b553b8545 |");
    expect(bodyOf(fake, 1453)).not.toContain("Legacy report");

    const after = snapshot(fake);
    await report(fake, floodFindings());
    expect(fake.edits.toSorted()).toEqual(canonical);
    expect(fake.created).toEqual([]);
    expect(snapshot(fake)).toBe(after);
  });
});

describe("a class issue closed as a duplicate", () => {
  const numberingFinding = (): Finding => {
    const failure = flood.failures.at(0);
    if (failure === undefined) throw new Error("Missing fixture");
    return failing(failure, 7);
  };
  const signature = [
    "consumer-flow",
    "plain",
    "tracked-changes",
    "Numbering provenance missing for <id>",
  ];
  const duplicate = (overrides: Partial<ClosedIssue> = {}): ClosedIssue => ({
    number: 1454,
    title: NUMBERING,
    body: oldMarker(signature),
    state: "closed",
    closedAt: "2026-10-02T09:00:00Z",
    disposition: SUPERSEDED,
    duplicateOf: 1300,
    ...overrides,
  });
  const target = (body = "Human notes"): Issue =>
    openIssue(1300, "Paragraph numbering is lost after a delete", body);

  test("reports to the issue it duplicates instead of reopening", async () => {
    const fake = fakeStore([target(), duplicate()]);
    const result = await report(fake, [numberingFinding()]);
    expect(result.map(({ issue }) => issue)).toEqual(["#1300"]);
    expect(fake.reopens).toEqual([]);
    expect(fake.created).toEqual([]);
    expect(fake.edits).toEqual([1300]);
    expect(bodyOf(fake, 1300)).toContain("Human notes");
    expect(rowCount(bodyOf(fake, 1300))).toBe(1);
    expect(bodyOf(fake, 1454)).toBe(oldMarker(signature));
    await report(fake, [numberingFinding()]);
    expect(writes(fake)).toBe(1);
  });

  test("follows a chain of duplicates; a loop keeps its issue closed", async () => {
    const chain = fakeStore([
      target(),
      duplicate(),
      duplicate({ number: 1457, duplicateOf: 1454 }),
      duplicate({ number: 1461, duplicateOf: 1457, disposition: { type: "duplicate" } }),
    ]);
    await report(chain, [numberingFinding()]);
    expect(chain.edits).toEqual([1300]);
    expect(chain.reopens).toEqual([]);
    expect(chain.created).toEqual([]);

    const loop = fakeStore([
      duplicate({ duplicateOf: 1457 }),
      duplicate({ number: 1457, duplicateOf: 1454 }),
    ]);
    await report(loop, [numberingFinding()]);
    expect(loop.reopens).toEqual([]);
    expect(loop.edits).toHaveLength(1);
    expect(loop.created).toEqual([]);
  });

  test("an unfixed closure with no duplicate target keeps rows and stays closed", async () => {
    const dispositions = [SUPERSEDED, { type: "duplicate" }] as const;
    for (const disposition of dispositions) {
      for (const closedAt of ["2026-10-02T09:00:00Z", "2026-01-01T00:00:00Z"]) {
        const fake = fakeStore([target(), duplicate({ duplicateOf: null, disposition, closedAt })]);
        const result = await report(fake, [numberingFinding()]);
        expect(result.map(({ issue }) => issue)).toEqual(["#1454"]);
        expect(fake.reopens).toEqual([]);
        expect(fake.created).toEqual([]);
        expect(fake.edits).toEqual([1454]);
        expect(fake.issues.find(({ number }) => number === 1454)?.state).toBe("closed");
        expect(rowCount(bodyOf(fake, 1454))).toBe(1);
        expect(bodyOf(fake, 1454)).toContain("do not reopen it");
        expect(bodyOf(fake, 1300)).toBe("Human notes");
        await report(fake, [numberingFinding()]);
        expect(writes(fake)).toBe(1);
      }
    }
  });

  test("a fixed closure reopens, even with a duplicate comment", async () => {
    const fake = fakeStore([target(), duplicate({ disposition: FIXED })]);
    const result = await report(fake, [numberingFinding()]);
    expect(result.map(({ issue }) => issue)).toEqual(["#1454"]);
    expect(fake.reopens).toEqual([1454]);
    expect(fake.edits).toEqual([1454]);
    expect(fake.created).toEqual([]);
    expect(bodyOf(fake, 1300)).toBe("Human notes");
  });

  test("a target it cannot write to is pointed at and nothing is written", async () => {
    const other = `<!-- fuzz-class: ${failureClass("consumer flow plain / direct", "no comment").key} -->`;
    for (const issues of [[duplicate()], [target(other), duplicate()]]) {
      const fake = fakeStore(issues);
      const result = await report(fake, [numberingFinding()]);
      expect(result.map(({ issue }) => issue)).toEqual(["#1300"]);
      expect(writes(fake)).toBe(0);
    }
  });

  test("a target closed long ago gets a linked new issue; the duplicate stays closed", async () => {
    const fake = fakeStore([
      { ...target(), state: "closed", closedAt: "2026-08-01T00:00:00Z", disposition: FIXED },
      duplicate(),
    ]);
    await report(fake, [numberingFinding()]);
    expect(fake.reopens).toEqual([]);
    expect(fake.edits).toEqual([]);
    expect(fake.created).toEqual([1455]);
    expect(bodyOf(fake, 1455)).toContain("Previous occurrence: #1300.");
  });

  test("the target is the issue the last duplicate comment names", () => {
    expect(duplicateTarget(["Duplicate of #1453"])).toBe(1453);
    expect(
      duplicateTarget(["Thanks.", "Duplicate of #12", "Also seen.\nduplicate of #34, see there"]),
    ).toBe(34);
    expect(duplicateTarget(["This is not a duplicate of #5", "See #7"])).toBeNull();
    expect(duplicateTarget([])).toBeNull();
  });
});

describe("the cap on new issues", () => {
  const names = ["alpha", "bravo", "charlie", "delta", "echo", "foxtrot"];
  const fresh = (count: number): Finding[] =>
    names
      .slice(0, count)
      .map((name, index) =>
        failing(
          { test: `consumer flow plain / direct`, message: `unexpected ${name} state` },
          index,
        ),
      );
  const standing = (): Issue[] => standingIssues().filter(({ number }) => number === 1389);
  const comment = (): Finding =>
    failing({ test: "consumer flow tables / direct", message: "no comment" }, 99);

  test("six new classes open nothing and fail; rows for known classes are still added", async () => {
    expect(MAX_NEW_ISSUES).toBe(5);
    const fake = fakeStore(standing());
    const error: unknown = await report(fake, [...fresh(3), comment(), ...fresh(6).slice(3)]).then(
      () => null,
      (reason: unknown) => reason,
    );
    expect(error).toBeInstanceOf(NewIssueCapError);
    const message = error instanceof Error ? error.message : "";
    expect(message).toContain("Refusing to open 6 fuzz failure issues in one run (limit 5)");
    for (const name of names) expect(message).toContain(`Consumer flow: unexpected ${name} state`);
    expect(fake.created).toEqual([]);
    expect(fake.issues).toHaveLength(1);
    expect(fake.edits).toEqual([1389]);
    expect(rowCount(bodyOf(fake, 1389))).toBe(1);
  });

  test("five new classes are opened, and only new classes count", async () => {
    const fake = fakeStore(standing());
    await report(fake, [...fresh(5), comment()]);
    expect(fake.created).toHaveLength(5);
    // Five classes now have issues: the sixth is the only new one.
    await report(fake, fresh(6));
    expect(fake.created).toHaveLength(6);
    expect(new Set(fake.issues.map(({ title }) => title)).size).toBe(7);
  });
});

test("issue evidence preserves comparison details beyond the former row cutoff", async () => {
  const fake = fakeStore();
  const item = finding(
    11,
    "expect(received).toEqual(expected)\n" +
      "context\n".repeat(100) +
      "- Expected field: before\n+ Received field: after",
  );
  await report(fake, [item]);
  const body = fake.issues.at(0)?.body;
  expect(body).toContain("- Expected field: before");
  expect(body).toContain("+ Received field: after");
});

describe("close reasons decide what a recurring class does", () => {
  const MESSAGES = [
    "listLevel is 3, expected 2",
    "Numbering provenance missing for 4f2a",
    "Cannot repack invalid DOCX document model",
  ] as const;
  const FIXTURES = ["lists", "plain", "tables"] as const;
  const MODES = ["direct", "tracked-changes"] as const;
  /** Numbers of legacy-path class issues that nightly runs reopened. */
  const CLASS_ISSUES = [1379, 1384, 1385] as const;
  const FOREIGN_ISSUE = 1300;
  const DISPOSITIONS = [
    { type: "fixed" },
    { type: "superseded" },
    { type: "duplicate" },
  ] as const satisfies readonly ClosedIssueDisposition[];
  const HOUR_MS = 60 * 60 * 1_000;
  const runContext = {
    ...context,
    runUrl: "https://github.com/stella/folio/actions/runs/18203344512",
    date: "2026-10-03T06:00:00Z",
  } satisfies Context;

  const findingArbitrary = fc.record({
    message: fc.constantFrom(...MESSAGES),
    fixture: fc.constantFrom(...FIXTURES),
    mode: fc.constantFrom(...MODES),
    seed: fc.integer({ min: 0, max: 2 ** 31 - 1 }),
    registered: fc.boolean(),
  });
  const issueArbitrary = fc.option(
    fc.oneof(
      fc.constant({ state: "open" } as const),
      fc.record({
        state: fc.constant("closed" as const),
        disposition: fc.constantFrom(...DISPOSITIONS),
        ageHours: fc.integer({ min: 0, max: 30 * 24 }),
        duplicateOf: fc.option(fc.constantFrom(...CLASS_ISSUES, FOREIGN_ISSUE)),
      }),
    ),
  );

  test(
    "superseded classes stay closed, reopens are explained, registered fingerprints write nothing",
    async () => {
      await assertProperty(
        fc.asyncProperty(
          fc.uniqueArray(findingArbitrary, {
            minLength: 1,
            maxLength: 8,
            selector: ({ message, fixture, mode }) => `${message}\0${fixture}\0${mode}`,
          }),
          fc.tuple(issueArbitrary, issueArbitrary, issueArbitrary),
          async (generated, issueShapes) => {
            const findings = generated.map(({ message, fixture, mode, seed }) =>
              failing({ test: `consumer flow ${fixture} / ${mode}`, message }, seed),
            );
            const known = new Map(
              generated.flatMap(({ registered }, index) => {
                const fingerprint = findings.at(index)?.record.marker.fingerprint;
                return registered && fingerprint !== undefined
                  ? [[fingerprint, `#${1500 + index}`] as const]
                  : [];
              }),
            );
            const unregistered = findings.filter(
              ({ record }) => !known.has(record.marker.fingerprint),
            );
            const unregisteredFingerprints = new Set(
              unregistered.map(({ record }) => record.marker.fingerprint),
            );
            const issues = issueShapes.flatMap((shape, index): Issue[] => {
              if (shape === null) return [];
              const message = MESSAGES[index] ?? MESSAGES[0];
              const sample = failing({ test: "consumer flow lists / direct", message }, 1);
              const { marker } = sample.record;
              const key = escaped(failureClass(marker.test, marker.assertion).key);
              const fields = {
                number: CLASS_ISSUES[index] ?? CLASS_ISSUES[0],
                title: `Consumer flow: ${message}`,
                body: `Owner notes\n<!-- fuzz-class: ${key} -->`,
              };
              if (shape.state === "open") return [{ ...fields, state: "open", closedAt: null }];
              const closedAt = Date.parse(runContext.date) - shape.ageHours * HOUR_MS;
              return [
                {
                  ...fields,
                  state: "closed",
                  closedAt: new Date(closedAt).toISOString(),
                  disposition: shape.disposition,
                  duplicateOf: shape.duplicateOf,
                },
              ];
            });

            const full = fakeStore(issues);
            await fileClasses({ findings, context: runContext, store: full.store, known });

            // (1) A class closed without a fix is never reopened.
            for (const issue of issues) {
              if (issue.state === "open" || issue.disposition.type === "fixed") continue;
              expect(full.reopens).not.toContain(issue.number);
              expect(full.issues.find(({ number }) => number === issue.number)?.state).toBe(
                "closed",
              );
            }
            // (2) Every reopen comments with the run and unregistered fingerprints.
            expect(full.comments.map(({ number }) => number)).toEqual(full.reopens);
            for (const { comment } of full.comments) {
              expect(comment).toContain(runContext.runUrl);
              const named = [...comment.matchAll(/`([0-9a-f]{16})`/gu)].map((match) => match[1]);
              expect(named.length).toBeGreaterThan(0);
              for (const fingerprint of named) {
                expect(unregisteredFingerprints.has(fingerprint ?? "")).toBe(true);
              }
            }
            // (3) Registered fingerprints change nothing: filing only the
            // unregistered findings writes exactly the same.
            const twin = fakeStore(issues);
            await fileClasses({
              findings: unregistered,
              context: runContext,
              store: twin.store,
              known,
            });
            expect(full.reopens).toEqual(twin.reopens);
            expect(full.comments).toEqual(twin.comments);
            expect(full.created).toEqual(twin.created);
            expect(full.edits).toEqual(twin.edits);
            expect(snapshot(full)).toBe(snapshot(twin));
            if (unregistered.length === 0) expect(writes(full)).toBe(0);
          },
        ),
      );
    },
    propertyTestTimeout(10_000),
  );
});
