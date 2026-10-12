import { expect, test } from "bun:test";
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

const workflow = Bun.YAML.parse(
  readFileSync(new URL("../.github/workflows/ci.yml", import.meta.url), "utf8"),
);
const plan = workflow.jobs["ci-plan"];
const lookup = plan.steps.find((step) => step.id === "reuse");
const warming = plan.steps.find((step) => step.id === "warm");
const pushedSha = "0123456789abcdef0123456789abcdef01234567";
const matchingRun = {
  id: 38238759782,
  event: "merge_group",
  status: "completed",
  conclusion: "success",
  head_sha: pushedSha,
};

const runLookup = (runs, status = 0) => {
  const directory = mkdtempSync(join(tmpdir(), "folio-push-reuse-"));
  const output = join(directory, "output");
  const requests = join(directory, "requests");
  writeFileSync(output, "");
  try {
    const result = Bun.spawnSync(
      [
        "bash",
        "-c",
        `gh() { printf '%s\\n' "$*" >> "$REQUESTS"; printf '%s' "$RESPONSE"; return "$STATUS"; }; ${lookup.run}`,
      ],
      {
        env: {
          ...process.env,
          RESPONSE: JSON.stringify({ workflow_runs: runs }),
          STATUS: String(status),
          REQUESTS: requests,
          REPO: "stella/folio",
          PUSHED_SHA: pushedSha,
          GITHUB_WORKFLOW_REF: "stella/folio/.github/workflows/ci.yml@refs/heads/main",
          GITHUB_OUTPUT: output,
        },
      },
    );
    return {
      status: result.exitCode,
      output: readFileSync(output, "utf8"),
      requests: readFileSync(requests, "utf8").trim().split("\n"),
    };
  } finally {
    rmSync(directory, { recursive: true, force: true });
  }
};

test("push coverage uses one exact-workflow API read and validates run identity", () => {
  expect(lookup.if).toBe("github.event_name == 'push'");
  const result = runLookup([matchingRun]);
  expect(result.status).toBe(0);
  expect(result.output).toBe("reused_run=38238759782\n");
  expect(result.requests).toEqual([
    `api --method GET repos/stella/folio/actions/workflows/ci.yml/runs -f event=merge_group -f status=success -f head_sha=${pushedSha} -f per_page=100`,
  ]);
});

test("missing, failed, malformed and mismatched coverage retains ordinary CI", () => {
  const cases = [
    [],
    [{ ...matchingRun, head_sha: "abcdef0123456789abcdef0123456789abcdef01" }],
    [{ ...matchingRun, event: "pull_request" }],
    [{ ...matchingRun, conclusion: "failure" }],
    [{ ...matchingRun, status: "in_progress" }],
    [{ ...matchingRun, id: "38238759782" }],
    [{ ...matchingRun, id: 0 }],
    null,
  ];
  for (const runs of cases) {
    const result = runLookup(runs);
    expect(result.status).toBe(0);
    expect(result.output).toBe("");
    expect(result.requests).toHaveLength(1);
  }
  const error = runLookup([matchingRun], 1);
  expect(error.status).toBe(0);
  expect(error.output).toBe("");
  expect(error.requests).toHaveLength(1);
});

test("warm-main only installs caches after reused coverage and dependency changes", () => {
  expect(warming.if).toBe("github.event_name == 'push' && steps.reuse.outputs.reused_run != ''");
  const warm = workflow.jobs["warm-main-caches"];
  expect(warm.if).toBe(
    "github.event_name == 'push' && needs.ci-plan.outputs.reused_run != '' && needs.ci-plan.outputs.warm_caches == 'true'",
  );
  expect(warm.steps.filter((step) => step.run).map((step) => step.run)).toEqual([
    "bun install --frozen-lockfile",
    "cargo install --git https://github.com/stella/tooling --rev a1bd81740ecb4b66999f56d1716e1190c4e7272f --locked stella_source_lint",
  ]);
  for (const step of warm.steps.filter(
    (candidate) =>
      candidate.id === "source-lint-cache" || candidate.name === "Install shared Rust source lint",
  )) {
    expect(workflow.jobs["docx-kernel"].steps).toContainEqual(step);
  }
});

test("the pushed range warms only for either lockfile or a toolchain change", () => {
  const directory = mkdtempSync(join(tmpdir(), "folio-push-range-"));
  const git = (...args) => {
    const result = Bun.spawnSync(["git", ...args], { cwd: directory });
    expect(result.exitCode).toBe(0);
    return result.stdout.toString().trim();
  };
  const commit = () => {
    git("add", ".");
    git(
      "-c",
      "user.name=Fixture",
      "-c",
      "user.email=fixture@example.test",
      "-c",
      "commit.gpgsign=false",
      "commit",
      "-qm",
      "chore: fixture",
    );
    return git("rev-parse", "HEAD");
  };
  try {
    git("init", "--quiet");
    writeFileSync(join(directory, "bun.lock"), "initial");
    writeFileSync(join(directory, "Cargo.lock"), "initial");
    writeFileSync(join(directory, "README.md"), "initial");
    let base = commit();
    for (const [file, toolchain, expected] of [
      ["README.md", "false", "false"],
      ["bun.lock", "false", "true"],
      ["Cargo.lock", "false", "true"],
      ["README.md", "true", "true"],
    ]) {
      writeFileSync(join(directory, file), `${file}-${base}`);
      const head = commit();
      const output = join(directory, "output");
      writeFileSync(output, "");
      const result = Bun.spawnSync(["bash", "-c", warming.run], {
        cwd: directory,
        env: {
          ...process.env,
          PUSH_BASE: base,
          PUSH_HEAD: head,
          TOOLCHAIN_CHANGED: toolchain,
          GITHUB_OUTPUT: output,
        },
      });
      expect(result.exitCode).toBe(0);
      expect(readFileSync(output, "utf8")).toBe(`changed=${expected}\n`);
      base = head;
    }
  } finally {
    rmSync(directory, { recursive: true, force: true });
  }
});
