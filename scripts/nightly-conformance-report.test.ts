import { expect, test } from "bun:test";
import type { ClosedIssueDisposition, Issue, IssueStore } from "./fuzz-issue-classes";
import { conformanceReportBody, fileConformanceReport } from "./nightly-failure-issues";
import { groupConformanceFailures, parseFailures, unparsedFailure } from "./nightly-failure-issues";

const context = {
  kind: "conformance" as const,
  runUrl: "https://github.com/stella/folio/actions/runs/12345678901",
  sha: "abcdef0123456789",
  date: "2026-10-02",
  factor: null,
};

const fakeStore = (initial: Issue[] = []) => {
  const issues = [...initial];
  const writes: string[] = [];
  const comments: string[] = [];
  const store = {
    list: async () => issues,
    create: async (title, body) => {
      const issue: Issue = {
        number: issues.length + 1,
        title,
        body,
        state: "open",
        closedAt: null,
      };
      issues.push(issue);
      writes.push("create");
      return issue;
    },
    edit: async (number, body) => {
      const issue = issues.find((candidate) => candidate.number === number);
      if (issue === undefined) throw new Error("Unknown test issue");
      issue.body = body;
      writes.push("edit");
    },
    reopen: async (number, comment) => {
      const issue = issues.find((candidate) => candidate.number === number);
      if (issue === undefined) throw new Error("Unknown test issue");
      const { title, body } = issue;
      issues[issues.indexOf(issue)] = { number, title, body, state: "open", closedAt: null };
      writes.push("reopen");
      comments.push(comment);
    },
  } satisfies IssueStore;
  return { issues, writes, comments, store };
};

const FIXED = { type: "fixed" } as const satisfies ClosedIssueDisposition;

test("arbitrary group counts and later runs update one standing issue, preserving legacy issues", async () => {
  const legacy = {
    number: 99,
    title: "Nightly conformance failure: standing report",
    state: "open",
    body: "Legacy evidence",
    closedAt: null,
  } as const satisfies Issue;
  const { store, issues, writes } = fakeStore([
    legacy,
    {
      number: 1404,
      title: "Conformance tracker",
      body: "Keep owner notes.\n<!-- standing-conformance -->\nKeep footer.",
      state: "open",
      closedAt: null,
    },
  ]);
  const failures = groupConformanceFailures(
    parseFailures(
      Array.from(
        { length: 200 },
        (_, index) =>
          `(fail) editor command conformance > paragraph › operation${index} @ document`,
      ).join("\n"),
    ),
  );
  expect(failures).toHaveLength(200);
  await fileConformanceReport({ failures, context, store });
  expect(issues).toHaveLength(2);
  expect(issues.at(0)).toEqual(legacy);
  expect(issues.at(1)?.number).toBe(1404);
  expect(issues.at(1)?.body).toContain("Keep owner notes.");
  expect(issues.at(1)?.body).toContain("Keep footer.");
  expect(issues.at(1)?.body).toContain("<!-- standing-conformance -->");
  expect(issues.at(1)?.body).toContain("Failure groups: 200.");
  expect(issues.at(1)?.body?.length).toBeLessThan(48_000);
  expect(issues.at(1)?.body).toContain("further failure groups omitted");
  await fileConformanceReport({ failures, context, store });
  expect(writes).toEqual(["edit"]);
  await fileConformanceReport({
    failures: [unparsedFailure("conformance")],
    context: { ...context, runUrl: "https://github.com/stella/folio/actions/runs/12345678902" },
    store,
  });
  expect(writes).toEqual(["edit", "edit"]);
  expect(issues).toHaveLength(2);
  expect(issues.at(1)?.body).toContain("crash or timeout");
  expect(issues.at(1)?.body).not.toContain("operation199");
});

test("a closed standing issue reopens even after fourteen days", async () => {
  const failures = [unparsedFailure("conformance")];
  const { store, writes, comments, issues } = fakeStore([
    {
      number: 100,
      title: "Nightly conformance failure: standing report",
      state: "closed",
      body: `<!-- standing-conformance -->\n\n${conformanceReportBody(failures, context)}`,
      closedAt: "2026-01-01T00:00:00Z",
      disposition: FIXED,
    },
  ]);
  await fileConformanceReport({ failures, context, store });
  expect(writes).toEqual(["reopen"]);
  expect(comments).toEqual([`Reopening: conformance failures recurred in ${context.runUrl}.`]);
  expect(issues.at(0)?.state).toBe("open");
  await fileConformanceReport({ failures, context, store });
  expect(writes).toEqual(["reopen"]);
});

test("marker matches anywhere in the body independently of the title", async () => {
  const failures = [unparsedFailure("conformance")];
  const body = `Owner notes\n<!-- standing-conformance -->\n\n${conformanceReportBody(failures, context)}`;
  const { store, writes, issues } = fakeStore([
    {
      number: 1,
      title: "Nightly conformance failure: standing report",
      state: "closed",
      body: "old",
      closedAt: "2026-10-01T00:00:00Z",
      disposition: FIXED,
    },
    { number: 2, title: "Renamed standing report", state: "open", body, closedAt: null },
  ]);
  await fileConformanceReport({ failures, context, store });
  expect(writes).toEqual([]);
  expect(issues.at(0)?.state).toBe("closed");
});

test("out-of-order runs never replace or reopen newer evidence", async () => {
  for (const state of ["open", "closed"] as const) {
    for (const olderRun of ["12345678900", "9999999999"]) {
      const failures = [unparsedFailure("conformance")];
      const body = `<!-- standing-conformance -->\n\n${conformanceReportBody(failures, context)}`;
      const { store, writes, issues } = fakeStore([
        state === "open"
          ? { number: 1404, title: "Tracker", body, state, closedAt: null }
          : { number: 1404, title: "Tracker", body, state, closedAt: null, disposition: FIXED },
      ]);
      await fileConformanceReport({
        failures: [],
        context: { ...context, runUrl: `https://github.com/stella/folio/actions/runs/${olderRun}` },
        store,
      });
      expect(writes).toEqual([]);
      expect(issues.at(0)?.body).toBe(body);
      expect(issues.at(0)?.state).toBe(state);
    }
  }
});

test("missing or ambiguous standing markers never create or update an issue", async () => {
  for (const count of [0, 2]) {
    const { store, writes } = fakeStore(
      Array.from(
        { length: count },
        (_, index) =>
          ({
            number: 1404 + index,
            title: "Tracker",
            state: "open",
            body: "<!-- standing-conformance -->",
            closedAt: null,
          }) satisfies Issue,
      ),
    );
    await expect(
      fileConformanceReport({ failures: [unparsedFailure("conformance")], context, store }),
    ).rejects.toThrow("Expected one issue marked");
    expect(writes).toEqual([]);
  }
});
