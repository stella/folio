const WORKSPACE_DEPENDENCIES = require("./scripts/workspace-dependency-policy.json");

const PHYSICAL_DEPENDENCY_TYPES = ["local"];

const workspaceNames = Object.keys(WORKSPACE_DEPENDENCIES);
const workspaceModules = Object.fromEntries(
  workspaceNames.map((workspace) => [
    workspace,
    require(`./packages/${workspace}/package.json`).name,
  ]),
);

const workspacePath = (workspace) => `^packages/${workspace}(?:/|$)`;
const workspaceTargetPattern = (workspaces) => {
  const packageDirectories = workspaces.join("|");
  const packageNames = workspaces
    .map((workspace) => workspaceModules[workspace].replaceAll("/", "\\/"))
    .join("|");
  return [`^packages/(?:${packageDirectories})(?:/|$)`, `^(?:${packageNames})(?:/|$)`];
};

const closedWorkspaceRules = Object.entries(WORKSPACE_DEPENDENCIES).flatMap(
  ([source, allowedTargets]) => {
    const otherWorkspaces = workspaceNames.filter((workspace) => workspace !== source);
    const forbiddenTargets = otherWorkspaces.filter(
      (workspace) => !allowedTargets.includes(workspace),
    );
    const rules = [
      {
        name: `${source}-uses-package-contracts`,
        comment:
          "Cross-package source access must use the target package name, never a relative path or TypeScript alias.",
        severity: "error",
        from: { path: workspacePath(source) },
        to: {
          path: `^packages/(?:${otherWorkspaces.join("|")})(?:/|$)`,
          dependencyTypes: PHYSICAL_DEPENDENCY_TYPES,
          dependencyTypesNot: ["aliased-workspace"],
        },
      },
    ];

    if (forbiddenTargets.length > 0) {
      rules.push({
        name: `${source}-workspace-dependencies`,
        comment: "Workspace packages may depend only on their explicitly owned lower layers.",
        severity: "error",
        from: { path: workspacePath(source) },
        to: { path: workspaceTargetPattern(forbiddenTargets) },
      });
    }

    return rules;
  },
);

module.exports = {
  forbidden: [...closedWorkspaceRules],
  options: {
    combinedDependencies: true,
    doNotFollow: {
      path: "node_modules",
      dependencyTypes: ["npm", "npm-dev", "npm-optional", "npm-peer", "npm-bundled", "npm-no-pkg"],
    },
    exclude: {
      path: ["(^|/)dist/", "(^|/)node_modules/", "^packages/docx-core/src/generated/"],
    },
    enhancedResolveOptions: {
      extensions: [".ts", ".tsx", ".mts", ".cts", ".js", ".jsx", ".mjs", ".cjs", ".vue"],
    },
    parser: "tsc",
    skipAnalysisNotInRules: true,
    tsConfig: {
      fileName: "tsconfig.depcruise.json",
    },
    tsPreCompilationDeps: true,
  },
};
