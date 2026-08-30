#!/usr/bin/env bun

import { existsSync, readFileSync, readdirSync } from "node:fs";
import path from "node:path";

const DEPENDENCY_FIELDS = [
  "dependencies",
  "devDependencies",
  "peerDependencies",
  "optionalDependencies",
] as const;
const ROOT = path.join(import.meta.dirname, "..");
const PACKAGES_DIRECTORY = path.join(ROOT, "packages");
const POLICY_PATH = path.join(import.meta.dirname, "workspace-dependency-policy.json");
const WORKSPACE_PROTOCOL = "workspace:";

type WorkspaceManifest = {
  directory: string;
  name: string;
  workspaceDependencies: ReadonlyArray<readonly [name: string, version: string]>;
};

type WorkspacePolicy = Map<string, Set<string>>;

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null;

const parseJson = (filePath: string): unknown => JSON.parse(readFileSync(filePath, "utf8"));

const readManifestDependencyEntries = (manifest: Record<string, unknown>) =>
  DEPENDENCY_FIELDS.flatMap((field) => {
    const dependencies = manifest[field];
    if (!isRecord(dependencies)) return [];
    return Object.entries(dependencies).flatMap(([name, version]) =>
      typeof version === "string" ? [[name, version] as const] : [],
    );
  });

const readWorkspaceManifests = (): WorkspaceManifest[] => {
  const rawManifests = readdirSync(PACKAGES_DIRECTORY, { withFileTypes: true }).flatMap((entry) => {
    const manifestPath = path.join(PACKAGES_DIRECTORY, entry.name, "package.json");
    if (!entry.isDirectory() || !existsSync(manifestPath)) return [];
    const manifest = parseJson(manifestPath);
    if (!isRecord(manifest) || typeof manifest["name"] !== "string") return [];
    return [{ directory: entry.name, manifest, name: manifest["name"] }];
  });
  const workspaceNames = new Set(rawManifests.map(({ name }) => name));

  return rawManifests.map(({ directory, manifest, name }) => ({
    directory,
    name,
    workspaceDependencies: readManifestDependencyEntries(manifest).filter(([dependency]) =>
      workspaceNames.has(dependency),
    ),
  }));
};

const readPolicy = (): WorkspacePolicy | null => {
  const rawPolicy = parseJson(POLICY_PATH);
  if (!isRecord(rawPolicy)) return null;

  const policy: WorkspacePolicy = new Map();
  for (const [source, targets] of Object.entries(rawPolicy)) {
    if (!Array.isArray(targets) || !targets.every((target) => typeof target === "string")) {
      return null;
    }
    policy.set(source, new Set(targets));
  }
  return policy;
};

const findCycle = (policy: WorkspacePolicy): string[] | null => {
  const visited = new Set<string>();
  const active = new Set<string>();
  const stack: string[] = [];

  const visit = (workspace: string): string[] | null => {
    if (active.has(workspace)) {
      const cycleStart = stack.indexOf(workspace);
      return [...stack.slice(cycleStart), workspace];
    }
    if (visited.has(workspace)) return null;

    active.add(workspace);
    stack.push(workspace);
    for (const target of policy.get(workspace) ?? []) {
      const cycle = visit(target);
      if (cycle !== null) return cycle;
    }
    stack.pop();
    active.delete(workspace);
    visited.add(workspace);
    return null;
  };

  for (const workspace of policy.keys()) {
    const cycle = visit(workspace);
    if (cycle !== null) return cycle;
  }
  return null;
};

export const validateWorkspaceManifestPolicy = (
  manifests: readonly WorkspaceManifest[],
  policy: WorkspacePolicy | null,
): string[] => {
  if (policy === null) return ["workspace dependency policy must be an object of string arrays"];

  const issues: string[] = [];
  const manifestDirectories = new Set(manifests.map(({ directory }) => directory));
  const packageToDirectory = new Map(manifests.map(({ directory, name }) => [name, directory]));

  for (const directory of manifestDirectories) {
    if (!policy.has(directory)) issues.push(`policy is missing packages/${directory}`);
  }
  for (const [source, targets] of policy) {
    if (!manifestDirectories.has(source)) {
      issues.push(`policy references missing packages/${source}`);
    }
    for (const target of targets) {
      if (!manifestDirectories.has(target)) {
        issues.push(`policy edge ${source} -> ${target} references a missing workspace`);
      }
      if (target === source) issues.push(`policy edge ${source} -> ${target} is self-referential`);
    }
  }

  for (const { directory, workspaceDependencies } of manifests) {
    const declaredTargets = new Set(
      workspaceDependencies.flatMap(([name]) => {
        const target = packageToDirectory.get(name);
        return target === undefined ? [] : [target];
      }),
    );
    const expectedTargets = policy.get(directory) ?? new Set();

    for (const target of declaredTargets) {
      if (!expectedTargets.has(target)) {
        issues.push(`packages/${directory} declares forbidden workspace dependency ${target}`);
      }
    }
    for (const target of expectedTargets) {
      if (!declaredTargets.has(target)) {
        issues.push(`packages/${directory} is missing declared workspace dependency ${target}`);
      }
    }
    for (const [name, version] of workspaceDependencies) {
      if (!version.startsWith(WORKSPACE_PROTOCOL)) {
        issues.push(`packages/${directory} dependency ${name} must use the workspace protocol`);
      }
    }
  }

  const cycle = findCycle(policy);
  if (cycle !== null) issues.push(`workspace dependency cycle: ${cycle.join(" -> ")}`);
  return issues.toSorted();
};

const main = () => {
  const issues = validateWorkspaceManifestPolicy(readWorkspaceManifests(), readPolicy());
  if (issues.length === 0) {
    process.stdout.write("Workspace manifest dependency graph is exact and acyclic.\n");
    return;
  }

  process.stderr.write(`${issues.map((issue) => `- ${issue}`).join("\n")}\n`);
  process.exitCode = 1;
};

if (import.meta.main) main();
