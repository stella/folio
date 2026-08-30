import { afterEach, describe, expect, setDefaultTimeout, test } from "bun:test";
import { mkdirSync, mkdtempSync, readFileSync, rmSync, symlinkSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

const REPOSITORY_ROOT = path.resolve(import.meta.dir, "..");
const DEPCRUISE_BINARY = path.join(REPOSITORY_ROOT, "node_modules/.bin/depcruise");
const DEPCRUISE_CONFIG = path.join(REPOSITORY_ROOT, ".dependency-cruiser.cjs");
const REPOSITORY_NODE_MODULES = path.join(REPOSITORY_ROOT, "node_modules");
const MANIFEST_CHECKER = path.join(REPOSITORY_ROOT, "scripts/check-workspace-manifests.ts");
const FIXTURE_POLICY = {
  agents: ["core"],
  core: ["docx-core"],
  "docx-core": [],
} as const;

const WORKSPACE_NAMES = {
  agents: "@stll/folio-agents",
  core: "@stll/folio-core",
  "docx-core": "@stll/docx-core",
} as const;

type Workspace = keyof typeof WORKSPACE_NAMES;

type WorkspaceDependency = {
  version?: string;
  workspace: Workspace;
};

type AddWorkspaceParams = {
  dependencies?: readonly WorkspaceDependency[];
  devDependencies?: readonly WorkspaceDependency[];
  fixtureRoot: string;
  workspace: Workspace;
};

type CruiseResult = {
  exitCode: number;
  output: string;
};

let fixtureRoots: string[] = [];

setDefaultTimeout(30_000);

afterEach(() => {
  for (const fixtureRoot of fixtureRoots) {
    rmSync(fixtureRoot, { force: true, recursive: true });
  }
  fixtureRoots = [];
});

const createFixture = (): string => {
  const fixtureRoot = mkdtempSync(path.join(tmpdir(), "folio-dependency-boundaries-"));
  fixtureRoots.push(fixtureRoot);
  mkdirSync(path.join(fixtureRoot, "node_modules/@stll"), { recursive: true });
  writeFileSync(
    path.join(fixtureRoot, "package.json"),
    JSON.stringify({ private: true, workspaces: ["packages/*"] }),
  );
  mkdirSync(path.join(fixtureRoot, "scripts"), { recursive: true });
  writeFileSync(
    path.join(fixtureRoot, "scripts/check-workspace-manifests.ts"),
    readFileSync(MANIFEST_CHECKER),
  );
  writeFileSync(
    path.join(fixtureRoot, "scripts/workspace-dependency-policy.json"),
    JSON.stringify(FIXTURE_POLICY),
  );
  writeFileSync(
    path.join(fixtureRoot, "tsconfig.base.json"),
    JSON.stringify({
      compilerOptions: {
        module: "Preserve",
        moduleResolution: "Bundler",
        noEmit: true,
        target: "ES2024",
      },
      include: ["packages"],
    }),
  );
  writeFileSync(
    path.join(fixtureRoot, "tsconfig.depcruise.json"),
    JSON.stringify({ extends: "./tsconfig.base.json", include: ["packages"] }),
  );
  return fixtureRoot;
};

const serializeDependencies = (dependencies: readonly WorkspaceDependency[]) =>
  Object.fromEntries(
    dependencies.map(({ version = "workspace:*", workspace }) => [
      WORKSPACE_NAMES[workspace],
      version,
    ]),
  );

const addWorkspace = ({
  dependencies = [],
  devDependencies = [],
  fixtureRoot,
  workspace,
}: AddWorkspaceParams): void => {
  const workspacePath = path.join(fixtureRoot, "packages", workspace);
  mkdirSync(path.join(workspacePath, "src"), { recursive: true });
  writeFileSync(
    path.join(workspacePath, "package.json"),
    JSON.stringify({
      dependencies: serializeDependencies(dependencies),
      devDependencies: serializeDependencies(devDependencies),
      exports: "./src/index.ts",
      main: "./src/index.ts",
      name: WORKSPACE_NAMES[workspace],
      type: "module",
      version: "0.0.0",
    }),
  );
  symlinkSync(
    path.join("..", "..", "packages", workspace),
    path.join(fixtureRoot, "node_modules/@stll", WORKSPACE_NAMES[workspace].slice("@stll/".length)),
    "dir",
  );
};

const writeSource = (
  fixtureRoot: string,
  workspace: Workspace,
  relativePath: string,
  source: string,
): void => {
  const sourcePath = path.join(fixtureRoot, "packages", workspace, "src", relativePath);
  mkdirSync(path.dirname(sourcePath), { recursive: true });
  writeFileSync(sourcePath, source);
};

const cruise = (fixtureRoot: string): CruiseResult => {
  const result = Bun.spawnSync(
    [DEPCRUISE_BINARY, "--config", DEPCRUISE_CONFIG, "--output-type", "err-long", "packages"],
    {
      cwd: fixtureRoot,
      env: { ...process.env, NODE_PATH: REPOSITORY_NODE_MODULES },
      stderr: "pipe",
      stdout: "pipe",
    },
  );
  return {
    exitCode: result.exitCode,
    output: `${result.stdout.toString()}${result.stderr.toString()}`,
  };
};

const checkManifests = (fixtureRoot: string): CruiseResult => {
  const result = Bun.spawnSync([process.execPath, "scripts/check-workspace-manifests.ts"], {
    cwd: fixtureRoot,
    stderr: "pipe",
    stdout: "pipe",
  });
  return {
    exitCode: result.exitCode,
    output: `${result.stdout.toString()}${result.stderr.toString()}`,
  };
};

describe("workspace dependency boundaries", () => {
  test("accepts an allowed declared package edge", () => {
    const fixtureRoot = createFixture();
    addWorkspace({ fixtureRoot, workspace: "docx-core" });
    addWorkspace({
      dependencies: [{ workspace: "docx-core" }],
      fixtureRoot,
      workspace: "core",
    });
    addWorkspace({ dependencies: [{ workspace: "core" }], fixtureRoot, workspace: "agents" });
    writeSource(fixtureRoot, "docx-core", "index.ts", "export const model = 1;\n");
    writeSource(
      fixtureRoot,
      "core",
      "index.ts",
      'import { model } from "@stll/docx-core";\nexport const document = model;\n',
    );

    const manifestResult = checkManifests(fixtureRoot);
    const cruiseResult = cruise(fixtureRoot);

    expect(manifestResult.exitCode, manifestResult.output).toBe(0);
    expect(cruiseResult.exitCode, cruiseResult.output).toBe(0);
  });

  test("rejects forbidden, undeclared, relative, require, and dynamic package edges", () => {
    const fixtureRoot = createFixture();
    addWorkspace({ fixtureRoot, workspace: "agents" });
    addWorkspace({
      dependencies: [{ workspace: "docx-core" }],
      fixtureRoot,
      workspace: "core",
    });
    addWorkspace({
      dependencies: [{ workspace: "core" }],
      fixtureRoot,
      workspace: "docx-core",
    });
    writeSource(fixtureRoot, "docx-core", "index.ts", "export const model = 1;\n");
    writeSource(
      fixtureRoot,
      "agents",
      "index.ts",
      'import { editor } from "@stll/folio-core";\nexport const tools = editor;\n',
    );
    writeSource(
      fixtureRoot,
      "core",
      "index.ts",
      'import { model } from "../../docx-core/src/index";\nexport const document = model;\n',
    );
    writeSource(
      fixtureRoot,
      "docx-core",
      "required.cjs",
      'const core = require("@stll/folio-core");\nmodule.exports = core;\n',
    );
    writeSource(
      fixtureRoot,
      "docx-core",
      "dynamic.mjs",
      'export const loadCore = () => import("@stll/folio-core");\n',
    );

    const manifestResult = checkManifests(fixtureRoot);
    const cruiseResult = cruise(fixtureRoot);

    expect(manifestResult.exitCode).not.toBe(0);
    expect(manifestResult.output).toContain(
      "packages/agents is missing declared workspace dependency core",
    );
    expect(manifestResult.output).toContain(
      "packages/docx-core declares forbidden workspace dependency core",
    );
    expect(cruiseResult.exitCode).not.toBe(0);
    expect(cruiseResult.output).toContain("core-uses-package-contracts");
    expect(cruiseResult.output).toContain("required.cjs");
    expect(cruiseResult.output).toContain("dynamic.mjs");
    expect(cruiseResult.output.match(/docx-core-workspace-dependencies/gu)).toHaveLength(2);
  });

  test("rejects a non-workspace version hidden by a duplicate declaration", () => {
    const fixtureRoot = createFixture();
    addWorkspace({ fixtureRoot, workspace: "docx-core" });
    addWorkspace({
      dependencies: [{ version: "0.0.0", workspace: "docx-core" }],
      devDependencies: [{ workspace: "docx-core" }],
      fixtureRoot,
      workspace: "core",
    });
    addWorkspace({ dependencies: [{ workspace: "core" }], fixtureRoot, workspace: "agents" });

    const result = checkManifests(fixtureRoot);

    expect(result.exitCode).not.toBe(0);
    expect(result.output).toContain(
      "packages/core dependency @stll/docx-core must use the workspace protocol",
    );
  });
});
