import { expect, setDefaultTimeout, test } from "bun:test";
import { readFileSync } from "node:fs";
import path from "node:path";

import { scanTestTimeout } from "../test/scan-timeout";
import {
  checkReusableWorkflowPermissions,
  loadPermissionWorkflow,
  parsePermissionWorkflow,
  REFRESH_REUSABLE_WORKFLOWS,
  REUSABLE_WORKFLOW_FIXTURES,
  workflowReference,
} from "./lib/reusable-workflow-permissions";

setDefaultTimeout(scanTestTimeout(30_000));

const SHA = "12eb83b8c6bd42a0bead73d9c6bc30b4caad6433";
const USES = `stella/.github/.github/workflows/release.yml@${SHA}`;
const caller = (permissions: unknown) =>
  parsePermissionWorkflow(
    Bun.YAML.stringify({
      permissions: { contents: "write", checks: "write" },
      jobs: { version: { uses: USES, permissions } },
    }),
  );
const callee = parsePermissionWorkflow(
  `permissions:\n  contents: read\n  pull-requests: read\njobs:\n  version:\n    permissions:\n      checks: read\n`,
);

test("every reusable workflow call covers its pinned callee permissions offline", () => {
  const root = path.resolve(import.meta.dir, "..");
  const used = new Set<string>();
  const problems: string[] = [];
  for (const file of new Bun.Glob(".github/workflows/*.{yml,yaml}").scanSync({
    cwd: root,
  })) {
    const workflow = parsePermissionWorkflow(
      readFileSync(path.join(root, file), "utf8"),
    );
    problems.push(
      ...checkReusableWorkflowPermissions({
        source: file,
        workflow,
        load: (reference) => {
          if (reference.type === "remote") used.add(reference.key);
          return loadPermissionWorkflow(root, reference);
        },
      }),
    );
  }
  expect(problems).toEqual([]);
  const vendored = new Set([
    ...new Bun.Glob("**/*.{yml,yaml}").scanSync({
      cwd: path.join(root, REUSABLE_WORKFLOW_FIXTURES),
      dot: true,
    }),
  ]);
  expect(
    vendored,
    `Refresh pinned workflows: ${REFRESH_REUSABLE_WORKFLOWS}`,
  ).toEqual(used);
});

test("missing root and job grants fail with the maintainer refresh command", () => {
  for (const missing of ["contents", "pull-requests", "checks"]) {
    const permissions = Object.fromEntries(
      ["contents", "pull-requests", "checks"]
        .filter((scope) => scope !== missing)
        .map((scope) => [scope, "read"]),
    );
    const problems = checkReusableWorkflowPermissions({
      source: "release-pr.yml",
      workflow: caller(permissions),
      load: () => callee,
    });
    expect(
      problems.some((problem) => problem.includes(`${missing}: read`)),
    ).toBe(true);
    expect(problems.at(-1)).toBe(
      `Refresh pinned workflows: ${REFRESH_REUSABLE_WORKFLOWS}`,
    );
  }
});

test("an unknown pin fails instead of reusing another pin fixture", () => {
  const unknown = "a".repeat(40);
  const workflow = parsePermissionWorkflow(
    `permissions: read-all\njobs:\n  version:\n    uses: stella/.github/.github/workflows/release.yml@${unknown}\n`,
  );
  const problems = checkReusableWorkflowPermissions({
    source: "release-pr.yml",
    workflow,
    load: (reference) =>
      reference.type === "remote" && reference.sha === SHA ? callee : undefined,
  });
  expect(problems.at(0)).toContain(
    `no pinned workflow fixture for stella/.github/${unknown}/release.yml`,
  );
  expect(problems.at(-1)).toContain(REFRESH_REUSABLE_WORKFLOWS);
});

test("job declarations replace root grants and write covers read", () => {
  const check = (permissions: unknown) =>
    checkReusableWorkflowPermissions({
      source: "release-pr.yml",
      workflow: caller(permissions),
      load: () => callee,
    });
  expect(check({ contents: "read", "pull-requests": "read" }).at(0)).toContain(
    "checks: read",
  );
  expect(
    check({ contents: "write", "pull-requests": "write", checks: "write" }),
  ).toEqual([]);
  expect(check("read-all")).toEqual([]);
  const inherited = caller(undefined);
  expect(
    checkReusableWorkflowPermissions({
      source: "release-pr.yml",
      workflow: inherited,
      load: () => callee,
    }).at(0),
  ).toContain("pull-requests: read");
});

test("nested calls preserve caller restrictions and require their own pinned fixtures", () => {
  const nested = parsePermissionWorkflow(
    `permissions:\n  contents: read\njobs:\n  child:\n    uses: ./.github/workflows/child.yml\n`,
  );
  const child = parsePermissionWorkflow(
    `permissions:\n  contents: write\njobs:\n  work:\n    runs-on: ubuntu-24.04\n`,
  );
  const problems = checkReusableWorkflowPermissions({
    source: "release-pr.yml",
    workflow: caller("write-all"),
    load: (reference) => (reference.file === "release.yml" ? nested : child),
  });
  expect(problems.at(0)).toContain("child.yml requires contents: write");
  const absent = checkReusableWorkflowPermissions({
    source: "release-pr.yml",
    workflow: caller("write-all"),
    load: (reference) =>
      reference.file === "release.yml" ? nested : undefined,
  });
  expect(absent.at(0)).toContain(`stella/.github/${SHA}/child.yml`);
});

test("local calls are resolved locally and mutable remote refs are rejected", () => {
  expect(workflowReference("./.github/workflows/local.yaml")).toEqual({
    type: "local",
    file: "local.yaml",
    key: "./.github/workflows/local.yaml",
  });
  expect(
    workflowReference("stella/.github/.github/workflows/release.yml@main"),
  ).toBeUndefined();
  const workflow = parsePermissionWorkflow(
    "permissions: {}\njobs:\n  call:\n    uses: stella/.github/.github/workflows/release.yml@main\n",
  );
  expect(
    checkReusableWorkflowPermissions({
      source: "release-pr.yml",
      workflow,
      load: () => undefined,
    }).at(0),
  ).toContain("full commit SHA");
});
