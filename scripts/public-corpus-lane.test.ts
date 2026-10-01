import { expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import path from "node:path";

const ROOT = path.resolve(import.meta.dir, "..");
const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);
const parseWorkflow = (source: string) => {
  const parsed: unknown = Bun.YAML.parse(source);
  if (!isRecord(parsed) || !isRecord(parsed["on"]) || !isRecord(parsed["jobs"]))
    throw new TypeError("Invalid workflow structure");
  // Bun folds duplicate keys; bind the block-style trigger declarations to its result.
  const section = /^on:\n([\s\S]*?)(?=^\S|$(?![\s\S]))/mu.exec(source)?.at(1);
  if (section === undefined) throw new TypeError("Missing block-style workflow triggers");
  const declarations = [...section.matchAll(/^  ([a-z_]+):/gmu)].map((match) => match[1]);
  if (new Set(declarations).size !== declarations.length)
    throw new TypeError("Duplicate workflow trigger declaration");
  expect(declarations.toSorted()).toEqual(Object.keys(parsed["on"]).toSorted());
  return { triggers: parsed["on"], jobs: parsed["jobs"] };
};
const workflow = (file: string) =>
  parseWorkflow(readFileSync(path.join(ROOT, ".github/workflows", file), "utf8"));

test("duplicate declarations of every configured trigger are rejected before parser folding", () => {
  const source = readFileSync(
    path.join(ROOT, ".github/workflows/nightly-public-corpus-flows.yml"),
    "utf8",
  );
  for (const trigger of Object.keys(parseWorkflow(source).triggers)) {
    const duplicate = source.replace("on:\n", `on:\n  ${trigger}:\n`);
    expect(() => parseWorkflow(duplicate)).toThrow("Duplicate workflow trigger declaration");
  }
});

// A reusable dispatcher must preserve both lanes after they share this workflow.
test("public corpus dispatch preserves the merged long-flow lane and avoids ordinary jobs", () => {
  const { triggers, jobs } = workflow("continuous-fuzz.yml");
  const dispatch = triggers["workflow_dispatch"];
  if (!isRecord(dispatch) || !isRecord(dispatch["inputs"]))
    throw new TypeError("Missing dispatcher inputs");
  const dedicated = Object.entries(jobs).filter(([, job]) => isRecord(job) && "uses" in job);
  expect(dedicated.map(([id]) => id).toSorted()).toEqual(["long-flows", "public-corpus-flows"]);
  for (const [id, job] of dedicated) {
    if (!isRecord(job)) throw new TypeError("Invalid reusable job");
    const input = id.replaceAll("-", "_");
    expect(dispatch["inputs"]).toHaveProperty(input);
    expect(job["if"]).toBe(`\${{ inputs.${input} }}`);
    const target = job["uses"];
    if (typeof target !== "string") throw new TypeError("Missing reusable workflow target");
    expect(workflow(path.basename(target)).triggers).toHaveProperty("workflow_call");
  }
  for (const id of ["consumer-fuzz", "property-fuzz"]) {
    const job = jobs[id];
    if (!isRecord(job)) throw new TypeError(`Missing ${id}`);
    expect(job["if"]).toBe("${{ !inputs.long_flows && !inputs.public_corpus_flows }}");
  }
});

test("public corpus generation stays outside PR gates and files findings with a separate token", () => {
  const { triggers, jobs } = workflow("nightly-public-corpus-flows.yml");
  expect(Object.keys(triggers).toSorted()).toEqual([
    "schedule",
    "workflow_call",
    "workflow_dispatch",
  ]);
  const fuzz = jobs["corpus-flows"];
  const report = jobs["report"];
  if (!isRecord(fuzz) || !isRecord(report)) throw new TypeError("Missing public corpus jobs");
  expect(fuzz["permissions"]).toEqual({ contents: "read" });
  expect(fuzz["timeout-minutes"]).toBe(45);
  expect(report["permissions"]).toEqual({ contents: "read", issues: "write" });
  expect(report["needs"]).toBe("corpus-flows");
  expect(report["timeout-minutes"]).toBe(5);
  expect(report["if"]).toBe("failure() && github.ref == 'refs/heads/main'");
  expect(report["concurrency"]).toEqual({
    group: "fuzz-failure-issues",
    "cancel-in-progress": false,
  });
  expect(JSON.stringify(report)).toContain("--log fuzz-results/run.log");
  expect(JSON.stringify(report)).toContain("--records fuzz-results");
  const dispatcher = workflow("continuous-fuzz.yml").jobs["public-corpus-flows"];
  if (!isRecord(dispatcher)) throw new TypeError("Missing reusable caller");
  expect(dispatcher["permissions"]).toEqual({ contents: "read", issues: "write" });
});
