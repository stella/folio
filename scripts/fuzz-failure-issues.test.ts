import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import { mkdtempSync, mkdirSync, readFileSync, writeFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { extractFailureMarkers } from "./failure-fingerprints";
import type { Issue } from "./fuzz-issue-classes";

import {
  failureMarker,
  failureRecord,
} from "../test/consumer-scenarios/support/failure-fingerprints";
import { assertProperty, propertyTestTimeout } from "../test/property-testing";
import {
  CONSUMER_SCENARIO_REPLAY_INPUTS,
  consumerScenarioInputs,
  parseConsumerArgs,
} from "./consumer-scenario-args";
import {
  AUTO_CLOSE_CLEAN_RUNS,
  type CloseReason,
  collectFindings,
  type Context,
  type Finding,
  type IssueContent,
  MAX_LISTED_SEEDS,
  parseCommentPages,
  parseFailureRecord,
  parseIssuePages,
  parseIssueResponse,
  parseSignatureState,
  fileFindings,
  readFindings,
  renderSignatureIssue,
  requireFindings,
  artifactInputs,
  RUN_TRIGGER,
  type SignatureIssueStore,
  type SignatureState,
  signatureOf,
  LABEL,
  TITLE_PREFIX,
  TITLE_FORMAT,
  cliRepro,
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
const RUN = "https://github.com/example/project/actions/runs/";
const runContext = (
  run: number,
  overrides: Partial<Pick<Context, "lane" | "trigger">> = {},
): Context => ({
  runUrl: `${RUN}${run}`,
  sha: null,
  source: "test",
  date: "",
  lane: overrides.lane ?? "nightly-swarm-fuzz",
  trigger: overrides.trigger ?? RUN_TRIGGER.scheduled,
});
const context = runContext(123);

const CLOSED_AS = {
  completed: { type: "fixed" },
  not_planned: { type: "superseded" },
  duplicate: { type: "duplicate" },
} as const satisfies Record<CloseReason, { type: string }>;

const fakeClient = (initial: Issue[] = []) => {
  const issues = [...initial];
  const creates: IssueContent[] = [];
  const edits: { number: number; content: IssueContent }[] = [];
  const reopens: number[] = [];
  const closes: { number: number; reason: CloseReason }[] = [];
  const replace = (number: number, update: (issue: Issue) => Issue): void => {
    const index = issues.findIndex((issue) => issue.number === number);
    const issue = issues.at(index);
    if (issue === undefined) throw new Error(`Unknown issue #${number}`);
    issues[index] = update(issue);
  };
  const store = {
    list: () => Promise.resolve(issues.map((issue) => ({ ...issue }))),
    create: (title, body) => {
      const issue = {
        number: issues.length + 1,
        title,
        body,
        state: "open",
        closedAt: null,
      } satisfies Issue;
      creates.push({ title, body });
      issues.push(issue);
      return Promise.resolve(issue);
    },
    edit: (number, content) => {
      edits.push({ number, content });
      replace(number, (issue) => ({ ...issue, ...content }));
      return Promise.resolve();
    },
    reopen: (number) => {
      reopens.push(number);
      replace(number, ({ title, body }) => ({
        number,
        title,
        body,
        state: "open",
        closedAt: null,
      }));
      return Promise.resolve();
    },
    close: (number, reason) => {
      closes.push({ number, reason });
      replace(number, ({ title, body }) => ({
        number,
        title,
        body,
        state: "closed",
        closedAt: "2026-10-17T00:00:00Z",
        disposition: CLOSED_AS[reason],
      }));
      return Promise.resolve();
    },
  } satisfies SignatureIssueStore;
  return { store, issues, creates, edits, reopens, closes };
};

const stateOf = (issue: Issue | undefined) => {
  if (issue === undefined) throw new Error("Missing issue");
  const state = parseSignatureState(issue);
  if (state === null) throw new Error(`Issue #${issue.number} has no signature state`);
  return state;
};

const flowFinding = (seed: number, flow: string, message = "step 3: no comment") => {
  const failure = new Error(message);
  const record = failureRecord(
    failureMarker({
      test: "consumer flow comments / suggested",
      seed,
      repro: `FOLIO_SCENARIO_SEED=${seed} bun scripts/consumer-scenarios.ts`,
      failure,
      flow,
    }),
    failure,
  );
  return { record, seeds: [seed], records: [record] } satisfies Finding;
};

test("a new signature opens one issue for every fingerprint and seed that share it", async () => {
  const findings = [
    flowFinding(1, "deleteBlock"),
    flowFinding(2, "deleteBlock+insertAfterBlock"),
    flowFinding(3, "commentOnBlock > accept all"),
  ];
  expect(new Set(findings.map(({ record }) => record.marker.fingerprint)).size).toBe(3);
  const fake = fakeClient();
  const { filed, closed } = await fileFindings({ findings, context, store: fake.store });
  expect(fake.creates).toHaveLength(1);
  expect(fake.edits).toHaveLength(0);
  expect(closed).toEqual([]);
  expect(filed.map(({ issue }) => issue)).toEqual(["#1", "#1", "#1"]);
  const first = findings.at(0);
  if (first === undefined) throw new Error("Missing finding");
  const signature = signatureOf(first.record.marker);
  expect(fake.creates.at(0)?.title).toBe(
    `Fuzz failure [${signature.id}]: consumer flow comments / suggested: no comment`,
  );
  const state = stateOf(fake.issues.at(0));
  expect(state.signature).toEqual(signature);
  expect(state.occurrences).toBe(3);
  expect(state.seeds).toEqual([1, 2, 3]);
  expect(state.lanes).toEqual([{ lane: "nightly-swarm-fuzz", lastRun: `${RUN}123`, cleanRuns: 0 }]);
  expect(state.lastSeen).toEqual({ lane: "nightly-swarm-fuzz", run: `${RUN}123` });
  expect(state.repro.command).toContain("--seed '1'");
  expect(state.repro.command).not.toContain("FOLIO_SCENARIO_");
});

test("signatures split by flow, mode and failure class, never by seed or label text", () => {
  const of = (name: string, assertion: string) => signatureOf({ test: name, assertion }).id;
  const flow = "consumer flow directNumbered / tracked-changes";
  expect(of(flow, "Missing numbering format for displayed label (<n>)")).toBe(
    of(flow, "Missing numbering format for displayed label •"),
  );
  expect(of(flow, "[batchSequential] the batch (direct: deleteBlock) leaves otherwise")).toBe(
    of(flow, "[batchSequential] the batch (direct: replaceBlock, insertTableRow) leaves otherwise"),
  );
  expect(of(flow, "block texts differ:")).not.toBe(of(flow, "comment"));
  expect(of(flow, "comment")).not.toBe(of("consumer flow directNumbered / direct", "comment"));
  expect(of(flow, "comment")).not.toBe(of("consumer flow lists / tracked-changes", "comment"));
  const fixtureSignatures = new Set(fixtureMarkers.map((item) => signatureOf(item).id));
  expect(fixtureSignatures.size).toBeLessThanOrEqual(
    new Set(fixtureMarkers.map(({ fingerprint }) => fingerprint)).size,
  );
});

test("fixture findings file one issue per signature, each exactly once", async () => {
  const fake = fakeClient();
  const { filed } = await fileFindings({ findings: fixtureFindings, context, store: fake.store });
  const signatures = new Set(fixtureMarkers.map((item) => signatureOf(item).id));
  expect(fake.creates).toHaveLength(signatures.size);
  expect(new Set(fake.issues.map((issue) => stateOf(issue).signature.id))).toEqual(signatures);
  expect(new Set(filed.map(({ fingerprint }) => fingerprint))).toEqual(
    new Set(fixtureMarkers.map(({ fingerprint }) => fingerprint)),
  );
  const again = await fileFindings({ findings: fixtureFindings, context, store: fake.store });
  expect(fake.creates).toHaveLength(signatures.size);
  expect(fake.edits).toHaveLength(0);
  expect(again.filed).toEqual(filed);
});

test("a repeat signature updates its issue in place with the run's seeds and count", async () => {
  const fake = fakeClient();
  await fileFindings({
    findings: [flowFinding(1, "deleteBlock"), flowFinding(2, "replaceBlock")],
    context,
    store: fake.store,
  });
  await fileFindings({
    findings: [flowFinding(2, "replaceBlock"), flowFinding(5, "insertAfterBlock")],
    context: runContext(124, { trigger: RUN_TRIGGER.manual, lane: "continuous-fuzz" }),
    store: fake.store,
  });
  expect(fake.creates).toHaveLength(1);
  expect(fake.edits.map(({ number }) => number)).toEqual([1]);
  expect(fake.reopens).toEqual([]);
  const state = stateOf(fake.issues.at(0));
  expect(state.occurrences).toBe(4);
  expect(state.seeds).toEqual([1, 2, 5]);
  expect(state.lastSeen).toEqual({ lane: "continuous-fuzz", run: `${RUN}124` });
  expect(state.lanes.map(({ lane }) => lane)).toEqual(["nightly-swarm-fuzz"]);
  await fileFindings({
    findings: [flowFinding(5, "insertAfterBlock")],
    context: runContext(124, { trigger: RUN_TRIGGER.manual, lane: "continuous-fuzz" }),
    store: fake.store,
  });
  expect(fake.edits).toHaveLength(1);
});

test("a fixed signature reopens when it recurs; one closed without a fix only records it", async () => {
  for (const [disposition, reopened] of [
    ["fixed", true],
    ["superseded", false],
    ["duplicate", false],
  ] as const) {
    const created = fakeClient();
    await fileFindings({ findings: [flowFinding(1, "a")], context, store: created.store });
    const issue = created.issues.at(0);
    if (issue === undefined) throw new Error("Missing issue");
    const fake = fakeClient([
      { ...issue, state: "closed", closedAt: "2026-10-10", disposition: { type: disposition } },
    ]);
    await fileFindings({
      findings: [flowFinding(9, "b")],
      context: runContext(200),
      store: fake.store,
    });
    expect(fake.creates).toHaveLength(0);
    expect(fake.reopens).toEqual(reopened ? [1] : []);
    expect(stateOf(fake.issues.at(0)).seeds).toEqual([1, 9]);
  }
});

test(`a signature closes after ${AUTO_CLOSE_CLEAN_RUNS} consecutive clean scheduled runs`, async () => {
  const fake = fakeClient();
  await fileFindings({ findings: [flowFinding(1, "a")], context, store: fake.store });
  const clean = (run: number, overrides: Partial<Pick<Context, "lane" | "trigger">> = {}) =>
    fileFindings({ findings: [], context: runContext(run, overrides), store: fake.store });
  for (let run = 1; run < AUTO_CLOSE_CLEAN_RUNS; run += 1) {
    expect((await clean(1000 + run)).closed).toEqual([]);
    await clean(1000 + run);
    await clean(2000 + run, { trigger: RUN_TRIGGER.manual });
    await clean(3000 + run, { lane: "nightly-property-test" });
  }
  expect(stateOf(fake.issues.at(0)).lanes.at(0)?.cleanRuns).toBe(AUTO_CLOSE_CLEAN_RUNS - 1);
  expect(fake.issues.at(0)?.state).toBe("open");
  await fileFindings({
    findings: [flowFinding(2, "b")],
    context: runContext(4000),
    store: fake.store,
  });
  expect(stateOf(fake.issues.at(0)).lanes.at(0)?.cleanRuns).toBe(0);
  for (let run = 1; run < AUTO_CLOSE_CLEAN_RUNS; run += 1) await clean(5000 + run);
  expect(fake.closes).toEqual([]);
  expect((await clean(6000)).closed).toEqual([1]);
  expect(fake.closes).toEqual([{ number: 1, reason: "completed" }]);
  expect(fake.issues.at(0)?.state).toBe("closed");
  await fileFindings({
    findings: [flowFinding(3, "c")],
    context: runContext(7000),
    store: fake.store,
  });
  expect(fake.reopens).toEqual([1]);
  expect(fake.creates).toHaveLength(1);
});

test("a signature seen by several lanes closes only once every lane ran clean", async () => {
  const fake = fakeClient();
  await fileFindings({ findings: [flowFinding(1, "a")], context, store: fake.store });
  const other = { lane: "nightly-property-test" };
  await fileFindings({
    findings: [flowFinding(2, "a")],
    context: runContext(124, other),
    store: fake.store,
  });
  const clean = (run: number, lane: string) =>
    fileFindings({ findings: [], context: runContext(run, { lane }), store: fake.store });
  for (let run = 0; run < AUTO_CLOSE_CLEAN_RUNS; run += 1)
    await clean(1000 + run, "nightly-swarm-fuzz");
  expect(fake.closes).toEqual([]);
  for (let run = 0; run < AUTO_CLOSE_CLEAN_RUNS; run += 1) await clean(2000 + run, other.lane);
  expect(fake.closes).toEqual([{ number: 1, reason: "completed" }]);
});

test("only scheduled runs register a lane; a merge-group repro never blocks closing", async () => {
  const fake = fakeClient();
  await fileFindings({
    findings: [flowFinding(1, "a")],
    context: runContext(1, { trigger: RUN_TRIGGER.mergeGroup, lane: "merge-queue-replay" }),
    store: fake.store,
  });
  expect(stateOf(fake.issues.at(0)).lanes).toEqual([]);
  for (let run = 0; run < AUTO_CLOSE_CLEAN_RUNS * 2; run += 1) {
    await fileFindings({ findings: [], context: runContext(100 + run), store: fake.store });
  }
  expect(fake.closes).toEqual([]);
  expect(fake.edits).toHaveLength(0);
});

const word = fc.stringMatching(/^[A-Za-z0-9<>()•.[\]/-]{1,8}$/u);
const sentence = fc.array(word, { minLength: 1, maxLength: 6 }).map((words) => words.join(" "));
const lane = fc.stringMatching(/^[a-z0-9][a-z0-9-]{0,12}$/u);
const run = fc.stringMatching(/^[a-z]{1,5}:[A-Za-z0-9/.-]{1,20}$/u);
const signatureState = fc
  .record({
    test: sentence,
    failure: sentence,
    occurrences: fc.nat(),
    seeds: fc.uniqueArray(fc.integer({ min: -(2 ** 31), max: 2 ** 31 }), {
      maxLength: MAX_LISTED_SEEDS,
    }),
    lastSeen: fc.record({ lane, run }),
    lanes: fc.uniqueArray(fc.record({ lane, lastRun: run, cleanRuns: fc.nat({ max: 99 }) }), {
      selector: ({ lane: name }) => name,
      maxLength: 4,
    }),
    assertion: fc.oneof(
      fc.string({ maxLength: 120 }),
      fc
        .tuple(lane, run, fc.nat({ max: 99 }), fc.string({ maxLength: 40 }))
        .map(([name, at, count, tail]) => `| ${name} | ${at} | ${count} |\n${tail}`),
    ),
    seed: fc.integer({ min: -(2 ** 31), max: 2 ** 31 }),
    command: fc.string({ minLength: 1, maxLength: 120 }),
  })
  .map(
    ({ test: name, failure, occurrences, seeds, lastSeen, lanes, assertion, seed, command }) =>
      ({
        signature: signatureOf({ test: name, assertion: failure }),
        occurrences,
        seeds,
        lastSeen,
        lanes,
        assertion,
        repro: { seed, command },
      }) satisfies SignatureState,
  );

test(
  "a signature issue body round-trips its state",
  () => {
    assertProperty(
      fc.property(signatureState, (state) => {
        const { body } = renderSignatureIssue(state);
        expect(parseSignatureState({ number: 1, body })).toEqual(state);
        expect(body.length).toBeLessThan(65_536);
      }),
      { numRuns: 200, id: "a signature issue body round-trips its state" },
    );
  },
  propertyTestTimeout(5_000),
);

test("bodies stay bounded for large evidence and seed sets", async () => {
  const original = fixtureMarkers.at(0);
  if (!original) throw new Error("Missing fixture marker");
  const markers = Array.from({ length: 1001 }, (_, seed) => ({
    ...original,
    seed,
    assertion: "`".repeat(100_000),
    repro: "`".repeat(100_000),
    test: "x".repeat(100_000),
  }));
  const fake = fakeClient();
  await fileFindings({
    findings: collectFindings([], markers),
    context: { ...context, runUrl: "x".repeat(100_000) },
    store: fake.store,
  });
  const issue = fake.issues.at(0);
  expect(issue?.body?.length).toBeLessThan(65_536);
  expect(issue?.title.length).toBeLessThanOrEqual(240);
  const state = stateOf(issue);
  expect(state.occurrences).toBe(1001);
  expect(state.seeds).toHaveLength(MAX_LISTED_SEEDS);
  expect(state.seeds.at(-1)).toBe(1000);
});

test("a 20 KB flow is emitted whole and its embedded JSON is unchanged", () => {
  const flow = {
    version: 1,
    steps: Array.from({ length: 800 }, (_, index) => ({ index, text: `step ${index}` })),
  };
  const json = JSON.stringify(flow);
  expect(json.length).toBeGreaterThan(20_000);
  const state: SignatureState = {
    signature: signatureOf({ test: "large flow", assertion: "failure" }),
    occurrences: 1,
    seeds: [41],
    lastSeen: { lane: "nightly", run: "local:test" },
    lanes: [],
    assertion: "failure",
    repro: {
      seed: 41,
      command: `bun scripts/consumer-scenarios.ts --seed 41 --flow '${json}' -- flow-corpus.test.ts`,
    },
  };
  const { body } = renderSignatureIssue(state);
  const embedded = /--flow '([^']+)'/u.exec(body)?.at(1);
  expect(JSON.parse(embedded ?? "null")).toEqual(flow);
  expect(parseSignatureState({ number: 1, body })?.repro.command).toBe(state.repro.command);
});

test("a huge assertion yields to the whole repro while the issue stays bounded", () => {
  const state: SignatureState = {
    signature: signatureOf({ test: "huge assertion", assertion: "failure" }),
    occurrences: 1,
    seeds: [42],
    lastSeen: { lane: "nightly", run: "local:test" },
    lanes: [],
    assertion: "assertion ".repeat(20_000),
    repro: {
      seed: 42,
      command: `bun scripts/consumer-scenarios.ts --seed 42 --flow '${"x".repeat(20_000)}'`,
    },
  };
  const { body } = renderSignatureIssue(state);
  expect(body.length).toBeLessThan(65_536);
  expect(body).toContain(state.repro.command);
  expect(body).toMatch(/… \(\d+ more characters\)/u);
});

test("an over-limit repro substitutes its flow but preserves every replay setting", () => {
  const state: SignatureState = {
    signature: signatureOf({ test: "oversized repro", assertion: "failure" }),
    occurrences: 1,
    seeds: [43],
    lastSeen: { lane: "nightly", run: "local:test" },
    lanes: [],
    assertion: "failure",
    repro: { seed: 43, command: "" },
  };
  const uniqueTail = "END-OF-OVERSIZED-REPRO";
  state.repro = {
    seed: 43,
    command: `bun scripts/consumer-scenarios.ts --seed 43 --relations 'preserve-comments' --relations-depth 3 --flow '${"x".repeat(70_000)}${uniqueTail}' -- flow-corpus.test.ts`,
  };
  const { body } = renderSignatureIssue(state);
  expect(body.length).toBeLessThan(65_536);
  expect(body).toContain(`findings/${state.signature.id}-43.flow.json`);
  expect(body).not.toContain(uniqueTail);
  expect(body).not.toContain("x".repeat(1_000));
  expect(body).toContain("--relations 'preserve-comments'");
  expect(body).toContain("--relations-depth 3");
  expect(body).toContain("-- flow-corpus.test.ts");
});

test("every legacy filer replay template is converted to CLI flags", () => {
  const templates = [
    "FOLIO_SCENARIO_SEED=1 FOLIO_SCENARIO_FUZZ_STEPS=2 FOLIO_SCENARIO_FUZZ_RUNS=3 FOLIO_SCENARIO_COLLISION_RUNS=4 FOLIO_SCENARIO_SWARM=1 bun scripts/consumer-scenarios.ts -- fuzz.test.ts",
    `FOLIO_SCENARIO_FLOW='{"steps":[]}' FOLIO_SCENARIO_PUBLIC_CORPUS_DIR='/tmp/corpus' bun scripts/consumer-scenarios.ts -- flow-corpus.test.ts`,
  ];
  for (const template of templates) expect(cliRepro(template)).not.toContain("FOLIO_SCENARIO_");
});

const replayArgv = (command: string): string[] =>
  [...command.matchAll(/'(?:'\\''|[^'])*'|\S+/gu)]
    .map(([token]) => (token.startsWith("'") ? token.slice(1, -1).replaceAll("'\\''", "'") : token))
    .slice(2);

test("legacy replay conversion merges environment and CLI inputs with CLI precedence", () => {
  const converted = cliRepro(
    "FOLIO_SCENARIO_SEED=1 FOLIO_SCENARIO_FUZZ_RUNS=3 bun scripts/consumer-scenarios.ts scenario.test.ts --seed 2 --only focused",
  );

  expect(converted.match(/--seed/gu)).toHaveLength(1);
  expect(converted.match(/--runs/gu)).toHaveLength(1);
  expect(parseConsumerArgs(replayArgv(converted), "/repo")).toMatchObject({
    seed: "2",
    runs: "3",
    only: "focused",
    files: ["scenario.test.ts"],
  });
});

test("legacy environment-only and flags-only replays resolve through the shared renderer", () => {
  const environmentOnly = cliRepro(
    "FOLIO_SCENARIO_SEED=1 bun scripts/consumer-scenarios.ts --only focused",
  );
  const flagsOnly = cliRepro("bun scripts/consumer-scenarios.ts --seed 2 --only focused");

  expect(parseConsumerArgs(replayArgv(environmentOnly), "/repo")).toMatchObject({
    seed: "1",
    only: "focused",
  });
  expect(parseConsumerArgs(replayArgv(flagsOnly), "/repo")).toMatchObject({
    seed: "2",
    only: "focused",
  });
});

test(
  "converted legacy replays resolve to the runner's merged inputs",
  () => {
    const replayInputs = CONSUMER_SCENARIO_REPLAY_INPUTS.map(({ env, flag }) => ({
      env,
      flag,
      option: flag.slice(2).replaceAll(/-([a-z])/gu, (_, letter: string) => letter.toUpperCase()),
    }));
    const options = replayInputs.map(({ option }) => option);
    const generatedInputs = fc.dictionary(
      fc.constantFrom(...options),
      fc.stringMatching(/^\/[a-zA-Z0-9_-]{1,20}$/u),
    );
    assertProperty(
      fc.property(generatedInputs, generatedInputs, (environmentValues, flagValues) => {
        const environment = Object.fromEntries(
          replayInputs.flatMap(({ env, option }) =>
            environmentValues[option] === undefined ? [] : [[env, environmentValues[option]]],
          ),
        );
        const environmentPrefix = Object.entries(environment)
          .map(([name, value]) => `${name}=${value}`)
          .join(" ");
        const legacyArgv = replayInputs.flatMap(({ flag, option }) =>
          flagValues[option] === undefined ? [] : [flag, flagValues[option]],
        );
        const legacy = `${environmentPrefix}${environmentPrefix === "" ? "" : " "}bun scripts/consumer-scenarios.ts --only focused ${legacyArgv.join(" ")}`;
        const converted = cliRepro(legacy);
        const resolvedBefore = consumerScenarioInputs(
          parseConsumerArgs(["--only", "focused", ...legacyArgv], "/repo"),
          environment,
        );
        const resolvedAfter = consumerScenarioInputs(
          parseConsumerArgs(replayArgv(converted), "/repo"),
          {},
        );

        expect(resolvedAfter).toEqual(resolvedBefore);
      }),
      { numRuns: 100, id: "converted legacy replay inputs preserve runner resolution" },
    );
  },
  propertyTestTimeout(5_000),
);

test("a malformed signature body fails instead of opening a duplicate", async () => {
  const fake = fakeClient([
    { number: 1, title: "x", body: "Signature: 0123456789abcdef\n", state: "open", closedAt: null },
  ]);
  await expect(fileFindings({ findings: [], context, store: fake.store })).rejects.toThrow(
    "Issue #1 has a malformed",
  );
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
      [
        "bun",
        "scripts/fuzz-failure-issues.ts",
        "--log",
        empty,
        "--lane",
        "nightly-swarm-fuzz",
        "--result",
        "failure",
        "--trigger",
        "scheduled",
        "--dry-run",
      ],
      { cwd: path.resolve(import.meta.dir, ".."), stdout: "pipe", stderr: "pipe" },
    );
    expect(cli.exitCode).not.toBe(0);
    expect(cli.stderr.toString()).toContain("0 FOLIO_FAILURE fingerprints");
  } finally {
    rmSync(root, { recursive: true, force: true });
  }
});

test("signature issue routing label and title format stay stable", () => {
  expect(LABEL.name).toBe("fuzz-failure");
  expect(TITLE_PREFIX).toBe("Fuzz failure");
  expect(TITLE_FORMAT).toBe("Fuzz failure [<signature>]: <test>: <failure>");
});
