import { describe, expect, test } from "bun:test";
import { mkdtempSync, mkdirSync, readFileSync, writeFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { extractFailureMarkers } from "./failure-fingerprints";
import type { Issue } from "./fuzz-issue-classes";

import {
  failureMarker,
  failureRecord,
} from "../test/consumer-scenarios/support/failure-fingerprints";
import {
  collectFindings,
  parseCommentPages,
  parseFailureRecord,
  parseIssuePages,
  parseIssueResponse,
  fileFindings,
  readFindings,
  requireFindings,
  artifactInputs,
  type FingerprintIssueStore,
  LABEL,
  TITLE_PREFIX,
  TITLE_FORMAT,
} from "./fuzz-failure-issues";

const marker = (seed: number, message = "step 3: no comment") =>
  failureMarker({
    test: "consumer flow comments / suggested",
    seed,
    repro: `FOLIO_SCENARIO_SEED=${seed} bun scripts/consumer-scenarios.ts --only '^fuzz run 0 \\('`,
    failure: new Error(message),
  });

describe("fuzz failure issues", () => {
  test("one finding per fingerprint; a record beats a bare marker and a shorter flow wins", () => {
    const long = failureRecord(marker(1), new Error("step 3: no comment"), {
      flow: { steps: [1, 2, 3] },
      shrink: { steps: 3, from: 10, attempts: 40 },
    });
    const short = failureRecord(marker(2), new Error("step 1: no comment"), {
      flow: { steps: [1] },
      shrink: { steps: 1, from: 12, attempts: 55 },
    });
    const other = marker(3, "block texts differ");
    const findings = collectFindings([long, short], [marker(1), marker(4), other]);
    expect(findings).toHaveLength(2);
    expect(findings[0]?.record).toBe(short);
    expect(findings[0]?.seeds).toEqual([1, 2, 4]);
    expect(findings[0]?.records?.map(({ marker: item }) => item.repro)).toEqual([
      marker(1).repro,
      marker(2).repro,
      marker(4).repro,
    ]);
    expect(findings[1]?.record.marker.fingerprint).toBe(other.fingerprint);
    expect(findings[1]?.record.replays).toEqual([other.repro]);
  });

  test("shrunk and unshrunk fingerprints retain separate issue identities", () => {
    const failure = new Error("step 3: no comment");
    const shrunk = failureRecord(
      failureMarker({
        test: "consumer flow comments / suggested",
        seed: 1,
        repro: "replay",
        failure,
        flow: "commentOnBlock > accept all",
      }),
      failure,
    );
    const findings = collectFindings([shrunk], [marker(5)]);
    expect(findings).toHaveLength(2);
    expect(findings[0]?.seeds).toEqual([1]);
    expect(findings[0]?.record.marker.primary).toBe(marker(5).fingerprint);
    expect(findings[0]?.records?.map(({ marker: item }) => item.seed)).toEqual([1]);
    expect(findings[1]?.seeds).toEqual([5]);
  });

  test("a shorter record also wins for the same seed's replay row", () => {
    const failure = new Error("step 3: no comment");
    const long = failureRecord(marker(1), failure, {
      replays: ["long replay"],
      shrink: { steps: 3, from: 10, attempts: 5 },
    });
    const short = failureRecord(marker(1), failure, {
      replays: ["short replay"],
      shrink: { steps: 1, from: 10, attempts: 6 },
    });
    const [finding] = collectFindings([long, short], [marker(1)]);
    expect(finding?.records).toEqual([short]);
  });

  test("records are validated before they reach an issue", () => {
    const record = failureRecord(marker(1), new Error("step 3: no comment"));
    expect(parseFailureRecord(JSON.parse(JSON.stringify(record)))).toEqual(record);
    expect(parseFailureRecord({ ...record, replays: ["one\nline"] })).toBeNull();
    expect(parseFailureRecord({ ...record, replays: [] })).toBeNull();
    expect(parseFailureRecord({ ...record, marker: { fingerprint: 1 } })).toBeNull();
    expect(parseFailureRecord({ ...record, version: 2 })).toBeNull();
  });
});

test("GitHub REST and CLI issue identities validate before filing", () => {
  const fields = { number: 20, title: "List numbering: listLevel mismatch", body: null };
  expect(parseIssueResponse({ ...fields, state: "open", closed_at: null })).toEqual({
    ...fields,
    state: "open",
    closedAt: null,
  });
  expect(parseIssueResponse({ ...fields, state: "CLOSED", closedAt: "2026-10-01" }).state).toBe(
    "closed",
  );
  expect(
    parseIssueResponse({
      ...fields,
      state: "closed",
      closed_at: "2026-10-01",
      state_reason: "not_planned",
    }),
  ).toEqual({
    ...fields,
    state: "closed",
    closedAt: "2026-10-01",
    disposition: { type: "superseded" },
  });
  // REST spells close reasons in lower case, the CLI's GraphQL fields in upper case.
  const dispositionOf = (stateReason: unknown) => {
    const issue = parseIssueResponse({
      ...fields,
      state: "CLOSED",
      closedAt: "2026-10-01",
      stateReason,
    });
    return issue.state === "closed" ? issue.disposition.type : null;
  };
  expect(dispositionOf("NOT_PLANNED")).toBe("superseded");
  expect(dispositionOf("COMPLETED")).toBe("fixed");
  expect(dispositionOf("completed")).toBe("fixed");
  expect(dispositionOf("DUPLICATE")).toBe("duplicate");
  // Issues closed before GitHub recorded close reasons were closed as done.
  expect(dispositionOf(null)).toBe("fixed");
  expect(dispositionOf("")).toBe("fixed");
  // A close reason the reporter does not know fails instead of reopening by default.
  expect(() => dispositionOf("reopened")).toThrow("Unknown close reason");
  expect(() => dispositionOf("wontfix")).toThrow("Unknown close reason");
  expect(() => dispositionOf(3)).toThrow("Unknown close reason");
  // An open issue's close reason ("reopened") does not apply.
  expect(
    parseIssueResponse({ ...fields, state: "open", closed_at: null, state_reason: "reopened" }),
  ).toEqual({ ...fields, state: "open", closedAt: null });
  expect(parseCommentPages([[{ body: "Duplicate of #7" }], [{ body: "later" }]])).toEqual([
    "Duplicate of #7",
    "later",
  ]);
  expect(() => parseCommentPages([[{ body: null }]])).toThrow();
  expect(() => parseCommentPages({})).toThrow();
  expect(() => parseIssueResponse({ ...fields, state: "unknown" })).toThrow();
  expect(() => parseIssueResponse({ ...fields, number: "20", state: "open" })).toThrow();
});

test("paginated issue responses fail before matching malformed issue data", () => {
  expect(parseIssuePages([])).toEqual([]);
  expect(() => parseIssuePages({})).toThrow();
  expect(() => parseIssuePages([{}])).toThrow();
  expect(() => parseIssuePages([[{ number: "20" }]])).toThrow();
  const issue = {
    number: 20,
    title: "Standing report",
    body: "Evidence",
    state: "open",
    closed_at: null,
  };
  expect(parseIssuePages([[issue], [issue]])).toHaveLength(2);
});

test("failure records retain the deepest comparison diff without terminal styling", () => {
  const comparison = new Error(
    "\u001b[31mexpect(received).toEqual(expected)\u001b[0m\n\n- Expected field: before\n+ Received field: after",
  );
  const wrapper = new Error("Generated flow failed", { cause: comparison });
  const record = failureRecord(marker(11, comparison.message), wrapper);
  expect(record.error).toContain("- Expected field: before");
  expect(record.error).toContain("+ Received field: after");
  expect(record.error).not.toContain("\u001b");
});

const fixtureLog = readFileSync(new URL("./fixtures/fuzz-failures.log", import.meta.url), "utf8");
const fixtureMarkers = extractFailureMarkers(fixtureLog);
const fixtureFindings = collectFindings([], fixtureMarkers);
const context = {
  runUrl: "https://github.com/example/project/actions/runs/123",
  sha: null,
  source: "test",
  date: "",
};

const fakeClient = (initial: Issue[] = []) => {
  const issues = [...initial];
  const creates: { title: string; body: string }[] = [];
  const reopens: number[] = [];
  const comments: { number: number; body: string }[] = [];
  const store = {
    list: () => Promise.resolve(issues),
    create: (title, body) => {
      const issue = {
        number: initial.length + creates.length + 1,
        title,
        body,
        state: "open",
        closedAt: null,
      } satisfies Issue;
      creates.push({ title, body });
      // fileFindings owns its fetched list; return the new issue without mutating it twice.
      return Promise.resolve(issue);
    },
    reopen: (number) => {
      reopens.push(number);
      return Promise.resolve();
    },
    comment: (number, body) => {
      comments.push({ number, body });
      return Promise.resolve();
    },
  } satisfies FingerprintIssueStore;
  return { store, creates, reopens, comments };
};

test("all 18 fixture fingerprints file exactly once, despite duplicate markers", async () => {
  expect(fixtureMarkers.length).toBeGreaterThan(18);
  expect(new Set(fixtureMarkers.map(({ fingerprint }) => fingerprint)).size).toBe(18);
  const fake = fakeClient();
  const filed = await fileFindings({ findings: fixtureFindings, context, store: fake.store });
  expect(fake.creates).toHaveLength(18);
  expect(fake.comments).toHaveLength(0);
  expect(fake.reopens).toHaveLength(0);
  expect(new Set(filed.map(({ issue }) => issue)).size).toBe(18);
  expect(new Set(filed.map(({ fingerprint }) => fingerprint))).toEqual(
    new Set(fixtureMarkers.map(({ fingerprint }) => fingerprint)),
  );
  for (const finding of fixtureFindings) {
    const created = fake.creates.find(({ title }) =>
      title.includes(finding.record.marker.fingerprint),
    );
    expect(created?.body).toContain(`Fingerprint: ${finding.record.marker.fingerprint}`);
    expect(created?.body).toContain(`Test: ${finding.record.marker.test}`);
    expect(created?.body).toContain(`Seed: ${finding.record.marker.seed}`);
    expect(created?.body).toContain(finding.record.marker.assertion);
    expect(created?.body).toContain(finding.record.marker.repro);
    expect(fake.creates.map(({ title }) => title)).toContain(
      `Fuzz failure [${finding.record.marker.fingerprint}]: ${finding.record.marker.test}`,
    );
  }
});

test("open and closed fixture issues receive exactly one run comment each; all closed reasons reopen", async () => {
  const initial = fixtureFindings.map(({ record: { marker: item } }, index) => {
    const disposition = (["superseded", "duplicate", "fixed"] as const).at(index % 3);
    if (disposition === undefined) throw new Error("Missing closed-issue fixture disposition");
    const fields = {
      number: index + 1,
      title: `Fuzz failure [${item.fingerprint}]: ${item.test}`,
      body: "x".repeat(100_000),
    };
    return index % 2 === 0
      ? ({ ...fields, state: "open", closedAt: null } satisfies Issue)
      : ({
          ...fields,
          state: "closed",
          closedAt: "2000-01-01",
          disposition: {
            type: disposition,
          },
        } satisfies Issue);
  });
  const fake = fakeClient(initial);
  // New seeds for the same fingerprints must reach the run comment, not another issue.
  const findings = collectFindings(
    [],
    [...fixtureMarkers, ...fixtureMarkers.map((item) => ({ ...item, seed: item.seed + 1000 }))],
  );
  const filed = await fileFindings({ findings, context, store: fake.store });
  expect(fake.creates).toHaveLength(0);
  expect(fake.reopens).toEqual(
    initial.filter(({ state }) => state === "closed").map(({ number }) => number),
  );
  expect(fake.comments).toHaveLength(18);
  expect(new Set(fake.comments.map(({ number }) => number)).size).toBe(18);
  expect(new Set(filed.map(({ issue }) => issue)).size).toBe(18);
  for (const [index, finding] of findings.entries()) {
    const comment = fake.comments.at(index)?.body;
    expect(comment).toContain(context.runUrl);
    for (const seed of finding.seeds) expect(comment).toContain(String(seed));
    expect(comment).toContain(finding.record.marker.fingerprint);
    expect(initial.at(index)?.body).toHaveLength(100_000);
  }
});

test("bodies and run comments remain bounded for large evidence and seed sets", async () => {
  const original = fixtureMarkers.at(0);
  if (!original) throw new Error("Missing fixture marker");
  const markers = Array.from({ length: 1001 }, (_, seed) => ({
    ...original,
    seed,
    assertion: "`".repeat(100_000),
    repro: "x".repeat(100_000),
  }));
  const findings = collectFindings([], markers);
  const fake = fakeClient([
    {
      number: 1,
      title: `Fuzz failure [${original.fingerprint}]: test`,
      body: "standing body",
      state: "open",
      closedAt: null,
    },
  ]);
  await fileFindings({ findings, context, store: fake.store });
  expect(fake.comments).toHaveLength(11);
  for (const { body } of fake.comments) expect(body.length).toBeLessThan(65_536);
  expect(fake.comments.at(-1)?.body).toContain("Seeds: 1000");
});

test("failed runs without markers fail; malformed marker lines carry a clear location", () => {
  expect(() =>
    requireFindings(collectFindings([], extractFailureMarkers("ordinary failure"))),
  ).toThrow("0 FOLIO_FAILURE fingerprints");
  expect(() => requireFindings(fixtureFindings)).not.toThrow();
  for (const line of [
    "FOLIO_FAILURE {bad",
    "FOLIO_FAILURE {}",
    "FOLIO_FAILURE",
    'FOLIO_FAILURE {"fingerprint":"wrong"}',
  ]) {
    expect(() => extractFailureMarkers(`ordinary log\n${line}`)).toThrow(
      "Malformed FOLIO_FAILURE line 2",
    );
  }
});

test("artifact discovery keeps matrix logs separate and scopes shared browser artifacts", () => {
  const root = mkdtempSync(path.join(tmpdir(), "folio-report-"));
  try {
    for (const directory of [
      "long-fuzz-200",
      "long-fuzz-1000",
      "browser/fuzz-artifacts/canonical",
      "browser/fuzz-artifacts/browser",
    ]) {
      mkdirSync(path.join(root, directory), { recursive: true });
      writeFileSync(path.join(root, directory, "run.log"), fixtureLog);
    }
    const extra = marker(42, "another mismatch");
    const recordDir = path.join(root, "long-fuzz-200/findings");
    mkdirSync(recordDir, { recursive: true });
    writeFileSync(
      path.join(recordDir, "failure.json"),
      JSON.stringify(failureRecord(extra, new Error("another mismatch"))),
    );
    const all = artifactInputs(root);
    expect(all.logs).toHaveLength(4);
    expect(readFindings(all.logs, all.records)).toHaveLength(19);
    const scoped = artifactInputs(root, "**/canonical/*.log");
    expect(scoped.logs).toHaveLength(1);
    expect(scoped.records).toEqual([path.join(root, "browser/fuzz-artifacts/canonical")]);
    const empty = path.join(root, "empty.log");
    writeFileSync(empty, "failed before emitting a marker");
    const cli = Bun.spawnSync(
      ["bun", "scripts/fuzz-failure-issues.ts", "--log", empty, "--require-failures", "--dry-run"],
      { cwd: path.resolve(import.meta.dir, ".."), stdout: "pipe", stderr: "pipe" },
    );
    expect(cli.exitCode).not.toBe(0);
    expect(cli.stderr.toString()).toContain("0 FOLIO_FAILURE fingerprints");
  } finally {
    rmSync(root, { recursive: true, force: true });
  }
});

test("fingerprint issue routing label and title format stay stable", () => {
  expect(LABEL.name).toBe("fuzz-failure");
  expect(TITLE_PREFIX).toBe("Fuzz failure");
  expect(TITLE_FORMAT).toBe("Fuzz failure [<fingerprint>]: <test>");
});
