import { afterEach, describe, expect, setDefaultTimeout, test } from "bun:test";
import {
  mkdirSync,
  mkdtempSync,
  readFileSync,
  rmSync,
  symlinkSync,
  writeFileSync,
} from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

import {
  expectExited,
  runChild,
  subprocessTestTimeout,
} from "../test/subprocess";

const REPOSITORY_ROOT = path.resolve(import.meta.dir, "..");
const DEPCRUISE_BINARY = path.join(
  REPOSITORY_ROOT,
  "node_modules/.bin/depcruise",
);
const DEPCRUISE_CONFIG = path.join(REPOSITORY_ROOT, ".dependency-cruiser.cjs");
const REPOSITORY_NODE_MODULES = path.join(REPOSITORY_ROOT, "node_modules");
const TYPESCRIPT_PRELOAD = path.join(
  REPOSITORY_ROOT,
  "scripts/lib/depcruise-typescript-preload.mjs",
);
const MANIFEST_CHECKER = path.join(
  REPOSITORY_ROOT,
  "scripts/check-workspace-manifests.ts",
);
const FIXTURE_POLICY = {
  agents: ["core"],
  core: ["docx-core"],
  "docx-core": [],
} as const;

const WORKSPACE_NAMES = {
  agents: "@stll/folio-agents",
  core: "@stll/folio-core",
  "docx-core": "@stll/docx-core",
  react: "@stll/folio-react",
  vue: "@stll/folio-vue",
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

setDefaultTimeout(subprocessTestTimeout(2));

afterEach(() => {
  for (const fixtureRoot of fixtureRoots) {
    rmSync(fixtureRoot, { force: true, recursive: true });
  }
  fixtureRoots = [];
});

const createFixture = (): string => {
  const fixtureRoot = mkdtempSync(
    path.join(tmpdir(), "folio-dependency-boundaries-"),
  );
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
    path.join(
      fixtureRoot,
      "node_modules/@stll",
      WORKSPACE_NAMES[workspace].slice("@stll/".length),
    ),
    "dir",
  );
};

const writeSource = (
  fixtureRoot: string,
  workspace: Workspace,
  relativePath: string,
  source: string,
): void => {
  const sourcePath = path.join(
    fixtureRoot,
    "packages",
    workspace,
    "src",
    relativePath,
  );
  mkdirSync(path.dirname(sourcePath), { recursive: true });
  writeFileSync(sourcePath, source);
};

const cruise = (fixtureRoot: string): CruiseResult => {
  const result = expectExited(
    runChild(
      [
        DEPCRUISE_BINARY,
        "--config",
        DEPCRUISE_CONFIG,
        "--output-type",
        "err-long",
        "packages",
      ],
      {
        cwd: fixtureRoot,
        env: {
          ...process.env,
          NODE_PATH: REPOSITORY_NODE_MODULES,
          NODE_OPTIONS: `--import=${TYPESCRIPT_PRELOAD}`,
        },
      },
    ),
  );
  return {
    exitCode: result.exitCode,
    output: `${result.stdout}${result.stderr}`,
  };
};

const checkManifests = (fixtureRoot: string): CruiseResult => {
  const result = expectExited(
    runChild([process.execPath, "scripts/check-workspace-manifests.ts"], {
      cwd: fixtureRoot,
    }),
  );
  return {
    exitCode: result.exitCode,
    output: `${result.stdout}${result.stderr}`,
  };
};

describe("workspace dependency boundaries", () => {
  test("keeps test support out of production editor source", () => {
    const fixtureRoot = createFixture();
    for (const workspace of ["core", "react", "vue"] as const) {
      addWorkspace({ fixtureRoot, workspace });
      writeSource(
        fixtureRoot,
        workspace,
        "editor/__tests__/helper.ts",
        "export const helper = 1;\nexport type Helper = number;\n",
      );
    }

    writeFileSync(
      path.join(fixtureRoot, "tsconfig.depcruise.json"),
      JSON.stringify({
        extends: "./tsconfig.base.json",
        compilerOptions: {
          baseUrl: ".",
          paths: { "@stll/folio-core/*": ["packages/core/src/*"] },
        },
        include: ["packages"],
      }),
    );
    writeSource(
      fixtureRoot,
      "core",
      "editor/index.ts",
      'export type { Helper } from "./__tests__/helper";\n',
    );
    writeSource(
      fixtureRoot,
      "react",
      "editor/index.tsx",
      'export const load = () => import("./__tests__/helper");\n',
    );
    writeSource(
      fixtureRoot,
      "vue",
      "editor/Editor.vue",
      '<script setup lang="ts">\nimport { helper } from "@stll/folio-core/editor/__tests__/helper";\n</script>\n<template>{{ helper }}</template>\n',
    );
    for (const source of [
      "editor/Editor.test.ts",
      "editor/Editor.spec.tsx",
      "editor/__fixtures__/fixture.ts",
      "editor/__tests__/consumer.ts",
    ]) {
      writeSource(
        fixtureRoot,
        "core",
        source,
        'export { helper } from "@stll/folio-core/editor/__tests__/helper";\n',
      );
    }

    const result = cruise(fixtureRoot);

    expect(result.exitCode, result.output).not.toBe(0);
    expect(
      result.output.match(
        /production-editor-source-does-not-import-test-support/gu,
      ),
      result.output,
    ).toHaveLength(3);
    for (const source of [
      "packages/core/src/editor/index.ts",
      "packages/react/src/editor/index.tsx",
      "packages/vue/src/editor/Editor.vue",
    ]) {
      expect(result.output).toContain(source);
    }
  });

  test("core review reaches the comments command path only through public re-exports or tests", () => {
    const fixtureRoot = createFixture();
    addWorkspace({ fixtureRoot, workspace: "core" });
    writeSource(
      fixtureRoot,
      "core",
      "prosemirror/commands/comments.ts",
      "export const review = 1;\n",
    );
    const forbiddenSources = {
      "ai-edits/headless.ts":
        'import { review } from "../prosemirror/commands/comments";\nexport const value = review;\n',
      "canonical/canonicalBatchReview.ts":
        'export const load = () => import("../prosemirror/commands/comments");\n',
      "controller/canonicalReview.ts":
        'export { review } from "../prosemirror/commands/comments";\n',
    };
    for (const [file, contents] of Object.entries(forbiddenSources)) {
      writeSource(fixtureRoot, "core", file, contents);
    }
    const allowedSources = {
      "prosemirror/commands/index.ts": "./comments",
      "prosemirror/commands/comments.test.ts": "./comments",
      "prosemirror/commands/__tests__/consumer.ts": "../comments",
    };
    for (const [file, specifier] of Object.entries(allowedSources)) {
      writeSource(
        fixtureRoot,
        "core",
        file,
        `export { review } from "${specifier}";\n`,
      );
    }

    const result = cruise(fixtureRoot);

    expect(result.exitCode, result.output).not.toBe(0);
    expect(
      result.output.match(/core-review-uses-dedicated-readers/gu),
      result.output,
    ).toHaveLength(Object.keys(forbiddenSources).length);
    for (const file of Object.keys(forbiddenSources)) {
      expect(result.output).toContain(`packages/core/src/${file}`);
    }
  });

  test("accepts an allowed declared package edge", () => {
    const fixtureRoot = createFixture();
    addWorkspace({ fixtureRoot, workspace: "docx-core" });
    addWorkspace({
      dependencies: [{ workspace: "docx-core" }],
      fixtureRoot,
      workspace: "core",
    });
    addWorkspace({
      dependencies: [{ workspace: "core" }],
      fixtureRoot,
      workspace: "agents",
    });
    writeSource(
      fixtureRoot,
      "docx-core",
      "index.ts",
      "export const model = 1;\n",
    );
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
    writeSource(
      fixtureRoot,
      "docx-core",
      "index.ts",
      "export const model = 1;\n",
    );
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
    expect(
      cruiseResult.output.match(/docx-core-workspace-dependencies/gu),
    ).toHaveLength(2);
  });

  test.each(["flat", "nested"] as const)(
    "rejects ProseMirror dependencies throughout docx-core, including type-only edges (%s)",
    (layout) => {
      const fixtureRoot = createFixture();
      addWorkspace({ fixtureRoot, workspace: "docx-core" });
      addWorkspace({ fixtureRoot, workspace: "core" });
      addWorkspace({ fixtureRoot, workspace: "agents" });
      const moduleRoot = path.join(
        fixtureRoot,
        "node_modules/prosemirror-model",
      );
      const frameworkRoot =
        layout === "flat"
          ? moduleRoot
          : path.join(
              fixtureRoot,
              "node_modules/.bun/prosemirror-model@0/node_modules/prosemirror-model",
            );
      mkdirSync(path.join(frameworkRoot, "dist"), { recursive: true });
      if (layout === "nested") symlinkSync(frameworkRoot, moduleRoot, "dir");
      writeFileSync(
        path.join(frameworkRoot, "package.json"),
        JSON.stringify({
          name: "prosemirror-model",
          main: "dist/index.js",
          types: "dist/index.d.ts",
        }),
      );
      writeFileSync(
        path.join(frameworkRoot, "dist/index.js"),
        "exports.Node = class Node {};\n",
      );
      writeFileSync(
        path.join(frameworkRoot, "dist/index.d.ts"),
        "export declare class Node {}\n",
      );
      const forbiddenSources = {
        "ops/static.ts":
          'import { Node } from "prosemirror-model";\nexport const node = Node;\n',
        "model/type.ts":
          'import type { Node } from "prosemirror-model";\nexport type EditorNode = Node;\n',
        "serialize/reexport.ts": 'export { Node } from "prosemirror-model";\n',
        "projection/dynamic.ts":
          'export const load = () => import("prosemirror-model");\n',
        "parse/required.cjs":
          'module.exports = require("prosemirror-model");\n',
      };
      for (const [relativePath, source] of Object.entries(forbiddenSources)) {
        writeSource(fixtureRoot, "docx-core", relativePath, source);
      }
      writeSource(
        fixtureRoot,
        "core",
        "editor.ts",
        'import { Node } from "prosemirror-model";\nexport const node = Node;\n',
      );

      const result = cruise(fixtureRoot);

      expect(result.exitCode, result.output).not.toBe(0);
      expect(
        result.output.match(/docx-core-stays-framework-free/gu),
      ).toHaveLength(Object.keys(forbiddenSources).length);
      for (const relativePath of Object.keys(forbiddenSources)) {
        expect(result.output).toContain(
          `packages/docx-core/src/${relativePath}`,
        );
      }
      expect(result.output).not.toContain("packages/core/src/editor.ts");
    },
  );

  test("rejects a non-workspace version hidden by a duplicate declaration", () => {
    const fixtureRoot = createFixture();
    addWorkspace({ fixtureRoot, workspace: "docx-core" });
    addWorkspace({
      dependencies: [{ version: "0.0.0", workspace: "docx-core" }],
      devDependencies: [{ workspace: "docx-core" }],
      fixtureRoot,
      workspace: "core",
    });
    addWorkspace({
      dependencies: [{ workspace: "core" }],
      fixtureRoot,
      workspace: "agents",
    });

    const result = checkManifests(fixtureRoot);

    expect(result.exitCode).not.toBe(0);
    expect(result.output).toContain(
      "packages/core dependency @stll/docx-core must use the workspace protocol",
    );
  });
});
