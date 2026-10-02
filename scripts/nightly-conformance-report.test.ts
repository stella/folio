import { expect, test } from "bun:test";
import type { Issue, IssueStore } from "./fuzz-issue-classes";
import {
  CONFORMANCE_REPORT_TITLE,
  conformanceReportBody,
  fileConformanceReport,
} from "./nightly-failure-issues";
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
    reopen: async (number) => {
      const issue = issues.find((candidate) => candidate.number === number);
      if (issue === undefined) throw new Error("Unknown test issue");
      issue.state = "open";
      writes.push("reopen");
    },
  } satisfies IssueStore;
  return { issues, writes, store };
};

test("arbitrary group counts and later runs update one standing issue, preserving legacy issues", async () => {
  const legacy = {
    number: 99,
    title: CONFORMANCE_REPORT_TITLE,
    state: "open",
    body: "Legacy evidence",
    closedAt: null,
  } as const satisfies Issue;
  const { store, issues, writes } = fakeStore([legacy]);
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
  expect(issues.at(1)?.title).toBe(CONFORMANCE_REPORT_TITLE);
  expect(issues.at(1)?.body).toContain("Failure groups: 200.");
  expect(issues.at(1)?.body?.length).toBeLessThan(48_000);
  expect(issues.at(1)?.body).toContain("further failure groups omitted");
  await fileConformanceReport({ failures, context, store });
  expect(writes).toEqual(["create"]);
  await fileConformanceReport({
    failures: [unparsedFailure("conformance")],
    context: { ...context, runUrl: "https://github.com/stella/folio/actions/runs/12345678902" },
    store,
  });
  expect(writes).toEqual(["create", "edit"]);
  expect(issues).toHaveLength(2);
  expect(issues.at(1)?.body).toContain("crash or timeout");
  expect(issues.at(1)?.body).not.toContain("operation199");
});

test("a closed standing issue reopens even after fourteen days", async () => {
  const failures = [unparsedFailure("conformance")];
  const { store, writes, issues } = fakeStore([
    {
      number: 100,
      title: CONFORMANCE_REPORT_TITLE,
      state: "closed",
      body: conformanceReportBody(failures, context),
      closedAt: "2026-01-01T00:00:00Z",
    },
  ]);
  await fileConformanceReport({ failures, context, store });
  expect(writes).toEqual(["reopen"]);
  expect(issues.at(0)?.state).toBe("open");
  await fileConformanceReport({ failures, context, store });
  expect(writes).toEqual(["reopen"]);
});

test("marker survives a title edit; an open standing issue takes precedence", async () => {
  const failures = [unparsedFailure("conformance")];
  const body = conformanceReportBody(failures, context);
  const { store, writes, issues } = fakeStore([
    {
      number: 1,
      title: CONFORMANCE_REPORT_TITLE,
      state: "closed",
      body: "old",
      closedAt: "2026-10-01T00:00:00Z",
    },
    { number: 2, title: "Renamed standing report", state: "open", body, closedAt: null },
  ]);
  await fileConformanceReport({ failures, context, store });
  expect(writes).toEqual([]);
  expect(issues.at(0)?.state).toBe("closed");
});
