import { panic } from "better-result";
import { existsSync, readFileSync } from "node:fs";
import path from "node:path";

export const REFRESH_REUSABLE_WORKFLOWS =
  "bun scripts/refresh-reusable-workflows.ts";
export const REUSABLE_WORKFLOW_FIXTURES = "test/workflow-contracts/reusable";

const ACCESS = { none: 0, read: 1, write: 2 } as const;
type Access = keyof typeof ACCESS;
type Permissions =
  | { type: "scoped"; grants: ReadonlyMap<string, Access> }
  | { type: "all"; access: "read" | "write" };

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const parsePermissions = (value: unknown): Permissions | undefined => {
  if (value === undefined) return undefined;
  if (value === "read-all") return { type: "all", access: "read" };
  if (value === "write-all") return { type: "all", access: "write" };
  if (!isRecord(value)) return panic("Invalid workflow permission declaration");
  const grants = new Map<string, Access>();
  for (const [scope, access] of Object.entries(value)) {
    if (access !== "none" && access !== "read" && access !== "write")
      return panic("Invalid workflow permission access", { scope, access });
    grants.set(scope, access);
  }
  return { type: "scoped", grants };
};

export const parsePermissionWorkflow = (source: string) => {
  const value: unknown = Bun.YAML.parse(source);
  if (!isRecord(value) || !isRecord(value["jobs"]))
    return panic("Invalid workflow jobs");
  const permissions = parsePermissions(value["permissions"]);
  const jobs = Object.entries(value["jobs"]).map(([name, job]) => {
    if (!isRecord(job)) return panic("Invalid workflow job", { name });
    const uses = job["uses"];
    if (uses !== undefined && typeof uses !== "string")
      return panic("Invalid reusable workflow call", { name });
    return { name, uses, permissions: parsePermissions(job["permissions"]) };
  });
  return { permissions, jobs };
};

export type WorkflowReference =
  | {
      type: "remote";
      repository: string;
      sha: string;
      file: string;
      key: string;
    }
  | { type: "local"; file: string; key: string };

export const workflowReference = (
  uses: string,
  parent?: WorkflowReference,
): WorkflowReference | undefined => {
  const local = /^\.\/\.github\/workflows\/(?<file>[\w.-]+\.ya?ml)$/u.exec(
    uses,
  )?.groups;
  if (local?.["file"] !== undefined) {
    const file = local["file"];
    if (parent?.type === "remote")
      return {
        type: "remote",
        repository: parent.repository,
        sha: parent.sha,
        file,
        key: `${parent.repository}/${parent.sha}/${file}`,
      };
    return { type: "local", file, key: `./.github/workflows/${file}` };
  }
  const remote =
    /^(?<repository>[\w.-]+\/[\w.-]+)\/\.github\/workflows\/(?<file>[\w.-]+\.ya?ml)@(?<sha>[a-f\d]{40})$/u.exec(
      uses,
    )?.groups;
  const repository = remote?.["repository"];
  const sha = remote?.["sha"];
  const file = remote?.["file"];
  if (repository === undefined || sha === undefined || file === undefined)
    return undefined;
  if (repository.split("/").some((part) => part === "." || part === ".."))
    return undefined;
  return {
    type: "remote",
    repository,
    sha,
    file,
    key: `${repository}/${sha}/${file}`,
  };
};

export const workflowFixturePath = (
  root: string,
  reference: WorkflowReference,
): string => {
  switch (reference.type) {
    case "local":
      return path.join(root, ".github/workflows", reference.file);
    case "remote":
      return path.join(root, REUSABLE_WORKFLOW_FIXTURES, reference.key);
    default: {
      const unreachable: never = reference;
      return unreachable;
    }
  }
};

export const loadPermissionWorkflow = (
  root: string,
  reference: WorkflowReference,
) => {
  const file = workflowFixturePath(root, reference);
  return existsSync(file)
    ? parsePermissionWorkflow(readFileSync(file, "utf8"))
    : undefined;
};

const gaps = (
  available: Permissions,
  required: Permissions | undefined,
): string[] => {
  if (required === undefined) return [];
  switch (required.type) {
    case "all":
      return available.type === "all" &&
        ACCESS[available.access] >= ACCESS[required.access]
        ? []
        : [`*: ${required.access}`];
    case "scoped":
      return Array.from(required.grants).flatMap(([scope, access]) => {
        const granted =
          available.type === "all"
            ? available.access
            : (available.grants.get(scope) ?? "none");
        return ACCESS[granted] >= ACCESS[access] ? [] : [`${scope}: ${access}`];
      });
    default: {
      const unreachable: never = required;
      return unreachable;
    }
  }
};

type PermissionCheckOptions = {
  source: string;
  workflow: ReturnType<typeof parsePermissionWorkflow>;
  load: (
    reference: WorkflowReference,
  ) => ReturnType<typeof parsePermissionWorkflow> | undefined;
};

/** Every call is checked against immutable callee declarations, entirely offline. */
export const checkReusableWorkflowPermissions = ({
  source,
  workflow,
  load,
}: PermissionCheckOptions): string[] => {
  const problems: string[] = [];
  const visit = ({
    current,
    available,
    parent,
    chain,
    label,
  }: {
    current: ReturnType<typeof parsePermissionWorkflow>;
    available?: Permissions;
    parent?: WorkflowReference;
    chain: ReadonlySet<string>;
    label: string;
  }): void => {
    for (const job of current.jobs) {
      if (job.uses === undefined) continue;
      const location = `${label}:${job.name}`;
      const reference = workflowReference(job.uses, parent);
      if (reference === undefined) {
        problems.push(
          `${location}: reusable workflow must use a full commit SHA: ${job.uses}`,
        );
        continue;
      }
      const callee = load(reference);
      if (callee === undefined) {
        problems.push(
          `${location}: no pinned workflow fixture for ${reference.key}`,
        );
        continue;
      }
      const granted = job.permissions ?? current.permissions ?? available;
      if (granted === undefined) {
        problems.push(`${location}: declare caller permissions explicitly`);
        continue;
      }
      const missing = new Set(gaps(granted, callee.permissions));
      for (const calleeJob of callee.jobs)
        for (const gap of gaps(
          granted,
          calleeJob.permissions ?? callee.permissions,
        ))
          missing.add(gap);
      if (missing.size > 0)
        problems.push(
          `${location}: ${reference.key} requires ${[...missing].join(", ")}`,
        );
      if (chain.has(reference.key)) {
        problems.push(
          `${location}: reusable workflow cycle at ${reference.key}`,
        );
        continue;
      }
      visit({
        current: callee,
        available: granted,
        parent: reference,
        chain: new Set([...chain, reference.key]),
        label: `${location} -> ${reference.key}`,
      });
    }
  };
  visit({ current: workflow, chain: new Set(), label: source });
  return problems.length === 0
    ? []
    : [...problems, `Refresh pinned workflows: ${REFRESH_REUSABLE_WORKFLOWS}`];
};
