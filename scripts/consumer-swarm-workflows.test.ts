import { expect, test } from "bun:test";
import { panic } from "better-result";
import { readFileSync } from "node:fs";

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const readWorkflow = (name: string) => {
  const source = readFileSync(new URL(`../.github/workflows/${name}`, import.meta.url), "utf8");
  const value: unknown = Bun.YAML.parse(source);
  if (!isRecord(value) || !isRecord(value["on"]) || !isRecord(value["jobs"])) {
    return panic(`${name}: invalid workflow`);
  }
  const triggerBlock = source.match(/^on:\n((?:[ \t].*\n|\n)*)/mu)?.[1];
  if (triggerBlock === undefined) return panic(`${name}: missing block triggers`);
  const declarations = [...triggerBlock.matchAll(/^  ([a-z_]+):$/gmu)].map((match) => match[1]);
  expect(declarations).toEqual(Object.keys(value["on"]));
  return { events: value["on"], jobs: value["jobs"] };
};

const job = (jobs: Record<string, unknown>, name: string) => {
  const value = jobs[name];
  if (!isRecord(value)) return panic(`missing job ${name}`);
  return value;
};

test("swarm workflows stay non-gating and preserve the long-flow dispatcher", () => {
  for (const name of ["continuous-fuzz.yml", "nightly-swarm-fuzz.yml"]) {
    const { events, jobs } = readWorkflow(name);
    expect(Object.keys(events).sort()).toEqual(["schedule", "workflow_dispatch"]);
    if (name !== "continuous-fuzz.yml") continue;
    const dispatch = events["workflow_dispatch"];
    if (!isRecord(dispatch) || !isRecord(dispatch["inputs"])) return panic("missing inputs");
    const inputs = dispatch["inputs"];
    const dedicated = Object.entries(jobs)
      .filter(([, value]) => isRecord(value) && typeof value["uses"] === "string")
      .map(([jobName]) => jobName.replaceAll("-", "_"))
      .sort();
    expect(Object.keys(inputs).sort()).toEqual(["factor", "minutes", "swarm", ...dedicated].sort());
    for (const jobName of ["consumer-fuzz", "property-fuzz"]) {
      expect(job(jobs, jobName)["if"]).toBe(
        `\${{ ${dedicated.map((input) => `!inputs.${input}`).join(" && ")} }}`,
      );
    }
    expect(job(jobs, "long-flows")["uses"]).toBe("./.github/workflows/nightly-long-fuzz.yml");
    expect(JSON.stringify(job(jobs, "consumer-fuzz"))).toContain("inputs.swarm");
  }
});

test("nightly swarm findings use a separate serialized main-only write token", () => {
  const { jobs } = readWorkflow("nightly-swarm-fuzz.yml");
  const execute = job(jobs, "swarm");
  expect(execute["permissions"]).toEqual({ contents: "read" });
  expect(execute["timeout-minutes"]).toBeLessThanOrEqual(45);
  const report = job(jobs, "report");
  expect(report["needs"]).toEqual(["swarm"]);
  expect(report["if"]).toBe("failure() && github.ref == 'refs/heads/main'");
  expect(report["permissions"]).toEqual({ contents: "read", issues: "write" });
  expect(report["concurrency"]).toEqual({
    group: "fuzz-failure-issues",
    "cancel-in-progress": false,
  });
  expect(JSON.stringify(report["steps"])).toContain("--records fuzz-results/fuzz-failures");
});
