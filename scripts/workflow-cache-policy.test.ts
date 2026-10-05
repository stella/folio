import { panic } from "better-result";
import { expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const events = (on: unknown): string[] => {
  if (typeof on === "string") return [on];
  if (Array.isArray(on)) return on.filter((event): event is string => typeof event === "string");
  return isRecord(on) ? Object.keys(on) : [];
};

const hasPublishToken = (workflow: unknown, job: unknown) => {
  if (!isRecord(job)) return false;
  const permissions =
    job["permissions"] ?? (isRecord(workflow) ? workflow["permissions"] : undefined);
  if (typeof permissions === "string") return permissions !== "read-all";
  return (
    isRecord(permissions) &&
    ["id-token", "contents", "packages"].some(
      (key) =>
        permissions[key] === "write" ||
        (typeof permissions[key] === "string" && permissions[key].includes("$" + "{{")),
    )
  );
};

const hasArtifactStep = (job: unknown, operation: string) =>
  isRecord(job) &&
  Array.isArray(job["steps"]) &&
  job["steps"].some(
    (step: unknown) =>
      isRecord(step) &&
      typeof step["uses"] === "string" &&
      step["uses"].startsWith(`actions/${operation}-artifact@`),
  );

// Protect the full dependency chain of a publishing token. Artifact readers
// can consume uploads without a needs edge, so include those producers too.
const publishingJobNames = (workflow: unknown) => {
  if (!isRecord(workflow) || !isRecord(workflow["jobs"])) return new Set<string>();
  const jobs = workflow["jobs"];
  const protectedJobs = new Set(
    Object.keys(jobs).filter((name) => hasPublishToken(workflow, jobs[name])),
  );
  const pending = [...protectedJobs];
  while (pending.length > 0) {
    const name = pending.pop();
    if (name === undefined) continue;
    const job = jobs[name];
    if (!isRecord(job)) continue;
    const needs = job["needs"];
    const dependencies = Array.isArray(needs)
      ? needs.filter((dependency): dependency is string => typeof dependency === "string")
      : [];
    if (typeof needs === "string") dependencies.push(needs);
    if (hasArtifactStep(job, "download")) {
      dependencies.push(
        ...Object.keys(jobs).filter((candidate) => hasArtifactStep(jobs[candidate], "upload")),
      );
    }
    if (hasArtifactStep(job, "upload")) {
      dependencies.push(
        ...Object.keys(jobs).filter((candidate) => hasArtifactStep(jobs[candidate], "download")),
      );
    }
    for (const dependency of dependencies) {
      if (protectedJobs.has(dependency)) continue;
      if (!isRecord(jobs[dependency])) continue;
      protectedJobs.add(dependency);
      pending.push(dependency);
    }
  }
  return protectedJobs;
};

const cacheProblems = (workflow: unknown): string[] => {
  if (!isRecord(workflow) || !isRecord(workflow["jobs"])) panic("Invalid workflow jobs");
  const defaultScope = events(workflow["on"]).some((event) =>
    ["workflow_run", "pull_request_target"].includes(event),
  );
  const publishing = publishingJobNames(workflow);
  return Object.entries(workflow["jobs"]).flatMap(([name, job]) => {
    if (!isRecord(job)) panic(`Invalid job ${name}`);
    if (typeof job["uses"] === "string") return [];
    if (!Array.isArray(job["steps"])) panic(`Missing steps in ${name}`);
    const steps = job["steps"].map((step: unknown) => {
      if (!isRecord(step)) panic(`Invalid step in ${name}`);
      return step;
    });
    const protectedInstall =
      defaultScope ||
      publishing.has(name) ||
      steps.some((step) => {
        const uses = step["uses"];
        return (
          uses === "./.github/actions/safe-chain" ||
          (typeof uses === "string" &&
            /(?:oven-sh\/setup-bun|stella\/\.github\/actions\/setup-bun-cached)@/u.test(uses) &&
            isRecord(step["with"]) &&
            step["with"]["no-cache"] === true)
        );
      });
    return steps.flatMap((step) => {
      const uses = typeof step["uses"] === "string" ? step["uses"] : "";
      const inputs = isRecord(step["with"]) ? step["with"] : {};
      const problems: string[] = [];
      if (!protectedInstall && uses.startsWith("oven-sh/setup-bun@")) {
        problems.push(`${name}: raw Bun setup needs the shared install-cache action`);
      }
      if (protectedInstall && uses.includes("setup-bun-cached@")) {
        problems.push(`${name}: shared install cache changes protected download policy`);
      }
      if (/^actions\/cache(?:\/[^@]+)?@/u.test(uses)) {
        const cachePath = inputs["path"];
        if (typeof cachePath === "string" && cachePath.includes(".bun/install/cache")) {
          problems.push(
            protectedInstall
              ? `${name}: install cache changes protected download policy`
              : `${name}: Bun install cache belongs to the shared action`,
          );
        }
      }
      if (uses.startsWith("rharkor/caching-for-turbo@") && inputs["server-port"] !== "0") {
        problems.push(`${name}: Turbo cache server needs server-port: "0"`);
      }
      return problems;
    });
  });
};

const raw = { uses: "oven-sh/setup-bun@fixture" };
const cached = { uses: "stella/.github/actions/setup-bun-cached@fixture" };

test("ordinary jobs use the cache owner and preserve protected runtime setup", () => {
  expect(cacheProblems({ jobs: { fixture: { steps: [raw] } } })).toHaveLength(1);
  expect(cacheProblems({ jobs: { fixture: { steps: [cached] } } })).toEqual([]);
  for (const on of ["workflow_run", ["pull_request_target"], { workflow_run: {} }]) {
    expect(cacheProblems({ on, jobs: { fixture: { steps: [raw] } } })).toEqual([]);
    expect(cacheProblems({ on, jobs: { fixture: { steps: [cached] } } })).toHaveLength(1);
  }
  for (const protection of [
    [{ uses: "./.github/actions/safe-chain" }],
    [{ ...raw, with: { "no-cache": true } }],
  ]) {
    expect(cacheProblems({ jobs: { fixture: { steps: [...protection, raw] } } })).toEqual([]);
    expect(cacheProblems({ jobs: { fixture: { steps: [...protection, cached] } } })).toHaveLength(
      1,
    );
  }
  expect(
    cacheProblems({ jobs: { fixture: { steps: [{ ...cached, with: { "no-cache": true } }] } } }),
  ).toHaveLength(1);
});

test("the cache owner replaces per-job download caches while other caches stay independent", () => {
  for (const uses of [
    "actions/cache@fixture",
    "actions/cache/restore@fixture",
    "actions/cache/save@fixture",
  ]) {
    expect(
      cacheProblems({
        jobs: { fixture: { steps: [{ uses, with: { path: "~/.bun/install/cache" } }] } },
      }),
    ).toHaveLength(1);
    expect(
      cacheProblems({ jobs: { fixture: { steps: [{ uses, with: { path: ".cache/corpus" } }] } } }),
    ).toEqual([]);
  }
});

test("Turbo cache servers request an available port", () => {
  const uses = "rharkor/caching-for-turbo@fixture";
  for (const port of [undefined, "41230", 0]) {
    expect(
      cacheProblems({ jobs: { fixture: { steps: [{ uses, with: { "server-port": port } }] } } }),
    ).toHaveLength(1);
  }
  expect(
    cacheProblems({ jobs: { fixture: { steps: [{ uses, with: { "server-port": "0" } }] } } }),
  ).toEqual([]);
});

test("every committed workflow and composite action follows the shared cache policy", () => {
  const root = fileURLToPath(new URL("../", import.meta.url));
  const workflows = [...new Bun.Glob(".github/workflows/*.{yml,yaml}").scanSync({ cwd: root })];
  expect(workflows.length).toBeGreaterThan(0);
  const problems = workflows.flatMap((file) =>
    cacheProblems(Bun.YAML.parse(readFileSync(`${root}/${file}`, "utf8"))).map(
      (problem) => `${file}: ${problem}`,
    ),
  );
  for (const file of new Bun.Glob(".github/actions/**/action.{yml,yaml}").scanSync({ cwd: root })) {
    const action: unknown = Bun.YAML.parse(readFileSync(`${root}/${file}`, "utf8"));
    if (!isRecord(action) || !isRecord(action["runs"])) panic(`Invalid action ${file}`);
    if (action["runs"]["using"] !== "composite") continue;
    problems.push(
      ...cacheProblems({ jobs: { composite: { steps: action["runs"]["steps"] } } }).map(
        (problem) => `${file}: ${problem}`,
      ),
    );
  }
  expect(problems).toEqual([]);
});

test("publishing tokens and their artifact chain reject cached Bun setup", () => {
  const rawSetup = {
    uses: "oven-sh/setup-bun@fixture",
    with: { "bun-version-file": "package.json" },
  };
  const cachedSetup = { ...rawSetup, uses: "stella/.github/actions/setup-bun-cached@fixture" };
  for (const permission of ["contents", "packages", "id-token"]) {
    expect(
      cacheProblems({
        jobs: { publish: { permissions: { [permission]: "write" }, steps: [cachedSetup] } },
      }),
    ).toHaveLength(1);
    expect(
      cacheProblems({
        jobs: { publish: { permissions: { [permission]: "write" }, steps: [rawSetup] } },
      }),
    ).toEqual([]);
  }
  expect(
    cacheProblems({ permissions: "write-all", jobs: { publish: { steps: [cachedSetup] } } }),
  ).toHaveLength(1);
  expect(
    cacheProblems({
      permissions: { contents: "write" },
      jobs: { ordinary: { permissions: { contents: "read" }, steps: [cachedSetup] } },
    }),
  ).toEqual([]);
  for (const needs of ["verify", ["verify"]]) {
    const workflow = {
      jobs: {
        build: { steps: [cachedSetup, { uses: "actions/upload-artifact@fixture" }] },
        verify: { needs: "build", steps: [] },
        publish: { permissions: { "id-token": "write" }, needs, steps: [] },
        ordinary: { steps: [cachedSetup] },
      },
    };
    expect(cacheProblems(workflow)).toHaveLength(1);
    workflow.jobs.build.steps[0] = rawSetup;
    expect(cacheProblems(workflow)).toEqual([]);
  }
  const consumers = {
    jobs: {
      build: {
        permissions: { contents: "write" },
        steps: [rawSetup, { uses: "actions/upload-artifact@fixture" }],
      },
      externalPublish: { steps: [cachedSetup, { uses: "actions/download-artifact@fixture" }] },
    },
  };
  expect(cacheProblems(consumers)).toHaveLength(1);
  const artifacts = {
    jobs: {
      build: { steps: [cachedSetup, { uses: "actions/upload-artifact@fixture" }] },
      publish: {
        permissions: { packages: "write" },
        steps: [{ uses: "actions/download-artifact@fixture" }],
      },
    },
  };
  expect(cacheProblems(artifacts)).toHaveLength(1);
});
