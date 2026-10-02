import { describe, expect, test } from "bun:test";
import {
  failureMarker,
  failureRecord,
} from "../test/consumer-scenarios/support/failure-fingerprints";
import { failureClass } from "./fuzz-failure-class";
import type { Context, Finding } from "./fuzz-failure-issues";
import { fileClasses, type Issue, type IssueStore } from "./fuzz-issue-classes";

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
    reopen: async (number: number) => {
      const issue = requireIssue(number);
      issue.state = "open";
      issue.closedAt = null;
      reopens.push(number);
    },
  } satisfies IssueStore;
  return { store, issues, edits, reopens, created };
};

const report = (fake: ReturnType<typeof fakeStore>, findings: Finding[], runContext = context) =>
  fileClasses({ findings, context: runContext, store: fake.store, known: new Map() });

const legacy = (closedAt: string | null = null): Issue => ({
  number: 20,
  title:
    "Fuzz failure [0123456789abcdef]: consumer flow lists / direct: listLevel is 3, expected 2",
  body: "Legacy failure details",
  state: closedAt === null ? "open" : "closed",
  closedAt,
});

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

  test("different classes create separate issues; creation cap never caps updates", async () => {
    const fake = fakeStore();
    await report(fake, [finding(1)]);
    const result = await fileClasses({
      findings: [finding(2), finding(3, "unrelated failure")],
      context,
      store: fake.store,
      known: new Map(),
      maxNewIssues: 0,
    });
    expect(fake.created).toEqual([1]);
    expect(fake.edits).toEqual([1]);
    expect(result.map(({ issue }) => issue)).toEqual(["#1", null]);
    await report(fake, [finding(3, "unrelated failure")]);
    expect(fake.created).toEqual([1, 2]);
  });

  test("recent closed class reopens with regression note, then rerun changes nothing", async () => {
    const fake = fakeStore([legacy("2026-09-18T12:00:01Z")]);
    await report(fake, [finding(1)]);
    expect(fake.reopens).toEqual([20]);
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

  test("any open legacy issue beats a closed marker issue; newest open wins", async () => {
    const signature = failureClass("consumer flow lists / direct", "listLevel is 3, expected 2");
    const marker = `<!-- fuzz-class: ${signature.key} -->`;
    const fake = fakeStore([
      legacy(),
      { ...legacy(), number: 21 },
      { ...legacy(context.date), number: 22, body: marker },
    ]);
    const result = await report(fake, [finding(1)]);
    expect(result.at(0)?.issue).toBe("#21");
    expect(fake.edits).toEqual([21]);
    expect(fake.reopens).toEqual([]);
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
    expect(fake.issues.at(0)?.body).toContain("| lists | direct |");
    await report(fake, [first]);
    expect(fake.edits).toEqual([20]);
  });

  test("split legacy aliases retain every issue regardless of record order", async () => {
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
      expect(new Set(result.map(({ issue }) => issue))).toEqual(new Set(["#20", "#21"]));
      expect(
        result.every(({ fingerprint }) => fingerprint === first.record.marker.fingerprint),
      ).toBe(true);
      expect(fake.created).toEqual([21]);
      expect(fake.edits).toEqual([20]);
    }
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
