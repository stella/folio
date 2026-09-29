import { describe, expect, test } from "bun:test";

import { hash32 } from "../test/commit-seed";
import { consumerSeed, judge, packageOf, propertyReplays } from "./fuzz-ejection";

const property = (fields: Record<string, unknown>): string =>
  `2026-09-29T10:00:00Z PROPERTY_FAILURE ${JSON.stringify({
    file: "packages/core/src/ai-edits/batchOverlap.test.ts",
    test: "refuses each conflict and applies the rest as one at a time would",
    seed: -440111207,
    path: "136:9:8:8",
    numRunsFactor: 1,
    ...fields,
  })}`;

describe("merge group fuzz replay", () => {
  test("the base replays the consumer flows under the seed the group's run derived", () => {
    const sha = "0302a42f5661aaaaaaaaaaaaaaaaaaaaaaaaaaaa";
    // scripts/consumer-scenarios.ts: commitSeed("consumer-scenarios fuzz") >>> 1.
    expect(consumerSeed(sha)).toBe(hash32(`${sha}\0consumer-scenarios fuzz`) >>> 1);
    expect(consumerSeed(sha)).toBeGreaterThanOrEqual(0);
  });

  test("property replays come from the failure lines, and only well-formed ones", () => {
    const log = [
      property({}),
      property({ path: "" }),
      property({ file: "../../etc/passwd.test.ts" }),
      property({ file: "packages/core/src/x.ts" }),
      property({ seed: "1; rm -rf /" }),
      property({ path: "1:2; echo" }),
      "PROPERTY_FAILURE {not json",
    ].join("\n");
    expect(propertyReplays(log)).toEqual([
      {
        file: "packages/core/src/ai-edits/batchOverlap.test.ts",
        title: "refuses each conflict and applies the rest as one at a time would",
        seed: -440111207,
        path: "136:9:8:8",
        factor: 1,
      },
      {
        file: "packages/core/src/ai-edits/batchOverlap.test.ts",
        title: "refuses each conflict and applies the rest as one at a time would",
        seed: -440111207,
        path: null,
        factor: 1,
      },
    ]);
    expect(packageOf("packages/core/src/ai-edits/batchOverlap.test.ts")).toEqual({
      dir: "packages/core",
      file: "src/ai-edits/batchOverlap.test.ts",
    });
    expect(packageOf("scripts/x.test.ts")).toEqual({ dir: ".", file: "scripts/x.test.ts" });
  });

  test("a group is pre-existing only when the base fails every way the group did", () => {
    const group = [
      { fingerprint: "aaaaaaaaaaaaaaaa", test: "consumer flow a / direct" },
      { fingerprint: "bbbbbbbbbbbbbbbb", test: "consumer flow b / direct" },
    ];
    expect(judge(group, new Set(["aaaaaaaaaaaaaaaa", "bbbbbbbbbbbbbbbb"]))).toMatchObject({
      verdict: "pre-existing",
      requeue: true,
    });
    expect(judge(group, new Set(["aaaaaaaaaaaaaaaa"]))).toMatchObject({
      verdict: "mixed",
      requeue: false,
    });
    expect(judge(group, new Set())).toMatchObject({ verdict: "introduced", requeue: false });
    expect(judge([], new Set(["aaaaaaaaaaaaaaaa"]))).toMatchObject({
      verdict: "none",
      requeue: false,
    });
  });
});
