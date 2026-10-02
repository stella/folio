import { describe, expect, test } from "bun:test";

import {
  failureMarker,
  failureRecord,
} from "../test/consumer-scenarios/support/failure-fingerprints";
import {
  collectFindings,
  parseFailureRecord,
  parseIssuePages,
  parseIssueResponse,
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

  test("an unshrunk report joins the shrunk finding of the same failure", () => {
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
    expect(findings).toHaveLength(1);
    expect(findings[0]?.seeds).toEqual([1, 5]);
    expect(findings[0]?.record.marker.primary).toBe(marker(5).fingerprint);
    expect(findings[0]?.records?.map(({ marker: item }) => item.seed)).toEqual([1, 5]);
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
