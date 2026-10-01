import { describe, expect, test } from "bun:test";

import {
  failureMarker,
  failureRecord,
} from "../test/consumer-scenarios/support/failure-fingerprints";
import {
  collectFindings,
  fingerprintOfTitle,
  issueBody,
  issueFingerprint,
  issueTitle,
  nextState,
  parseFailureRecord,
  readState,
} from "./fuzz-failure-issues";

const marker = (seed: number, message = "step 3: no comment") =>
  failureMarker({
    test: "consumer flow comments / suggested",
    seed,
    repro: `FOLIO_SCENARIO_SEED=${seed} bun scripts/consumer-scenarios.ts --only '^fuzz run 0 \\('`,
    failure: new Error(message),
  });

const context = {
  runUrl: "https://github.com/stella/folio/actions/runs/1",
  sha: "0302a42f5661aaaaaaaaaaaaaaaaaaaaaaaaaaaa",
  source: "continuous fuzz",
  date: "2026-09-29",
};

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
    expect(nextState(findings[0]!, null, "2026-09-29").primary).toBe(marker(5).fingerprint);
  });

  test("the title carries the fingerprint the next run finds it by", () => {
    const [finding] = collectFindings([], [marker(1)]);
    const title = issueTitle(finding!);
    expect(title).toBe(
      `Fuzz failure [${finding!.record.marker.fingerprint}]: consumer flow comments / suggested: no comment`,
    );
    expect(fingerprintOfTitle(title)).toBe(finding!.record.marker.fingerprint);
    expect(fingerprintOfTitle("Nightly property failure: x")).toBeNull();
    const [long] = collectFindings([], [marker(1, `step 1: ${"x".repeat(400)}`)]);
    expect(issueTitle(long!).length).toBeLessThanOrEqual(240);
  });

  test("terse manually filed titles deduplicate through the existing body state", () => {
    const [finding] = collectFindings([], [marker(7)]);
    if (!finding) throw new Error("fixture has no finding");
    const body = issueBody(finding, nextState(finding, null, context.date), context);
    expect(
      issueFingerprint({ title: "Document operations: formatting differs after insertion", body }),
    ).toBe(finding.record.marker.fingerprint);
    expect(issueFingerprint({ title: issueTitle(finding), body: null })).toBe(
      finding.record.marker.fingerprint,
    );
    expect(issueFingerprint({ title: "Other issue", body: null })).toBeNull();
  });

  test("a recurrence updates the count and seeds kept in the body", () => {
    const [finding] = collectFindings(
      [
        failureRecord(marker(7), new Error("step 3: no comment"), {
          replays: ["FOLIO_SCENARIO_FLOW='{}' bun scripts/consumer-scenarios.ts", marker(7).repro],
          flow: { version: 1 },
          shrink: { steps: 2, from: 10, attempts: 31 },
        }),
      ],
      [],
    );
    const first = issueBody(finding!, nextState(finding!, null, "2026-09-28"), context);
    expect(first).toContain("### Replay\n```sh\nFOLIO_SCENARIO_FLOW='{}'");
    expect(first).toContain("Also replays with:");
    expect(first).toContain("### Minimized flow (2 steps, shrunk from 10 in 31 replays)");
    expect(first).toContain("Seen in 1 run since 2026-09-28");
    const again = nextState({ ...finding!, seeds: [9] }, readState(first), context.date);
    expect(again).toMatchObject({ count: 2, firstSeen: "2026-09-28", lastSeen: "2026-09-29" });
    expect(again.seeds).toEqual([9, 7]);
    expect(issueBody(finding!, again, context)).toContain("Seen in 2 runs since 2026-09-28");
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
