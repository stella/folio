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

const coreSource = (directory) => `^packages/core/src/${directory}(?:/|$)`;

/**
 * Modules a paint backend must not see. A backend that can read the layout
 * engine can quietly grow a second opinion about where something goes, which
 * is the exact divergence the display list exists to prevent: whatever a
 * backend needs has to be added to the display list, where every backend sees
 * it, rather than fetched behind the seam.
 */
const LAYOUT_OWNED_SOURCES = [
  "layout-engine",
  "layout-painter",
  "layout-bridge",
  "paged-layout",
  "prosemirror",
  "controller",
  "docx",
  "ai-edits",
  "compare",
  // The producer knows about layout by design, so reaching it is reaching
  // layout one step removed.
  "display-list/build",
].map(coreSource);

const paintBackendRules = [
  {
    name: "paint-backends-read-only-the-display-list",
    comment:
      "The PDF and display-list DOM backends may not reach into layout: a fact a backend needs belongs in the display list, where both backends see it.",
    severity: "error",
    from: { path: [coreSource("pdf"), coreSource("display-list/dom")] },
    to: { path: LAYOUT_OWNED_SOURCES, dependencyTypes: PHYSICAL_DEPENDENCY_TYPES },
  },
  {
    name: "display-list-types-stay-pure-data",
    comment: "The paint IR is serializable data: it may not import anything from core.",
    severity: "error",
    from: { path: "^packages/core/src/display-list/types\\.ts$" },
    to: { path: "^packages/core/src/", dependencyTypes: PHYSICAL_DEPENDENCY_TYPES },
  },
];

module.exports = {
  forbidden: [...closedWorkspaceRules, ...paintBackendRules],
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
