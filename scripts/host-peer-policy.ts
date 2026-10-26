import { TaggedError } from "better-result";
import { existsSync, readdirSync, readFileSync, realpathSync } from "node:fs";
import { createRequire } from "node:module";
import path from "node:path";

import { parseSourceModule, walkAst } from "./lib/oxc-ast";

const SINGLETON =
  /^(?:prosemirror-[a-z-]+|yjs|y-prosemirror|y-protocols|react|react-dom)(?:\/|$)/u;
const TEST_SOURCE =
  /(?:^|\/)(?:__tests__|test|tests|fixtures|__fixtures__|__snapshots__)(?:\/|$)|\.(?:test|spec)\./u;
const BUNDLED_SENTINELS = [
  "react-stack-bottom-frame",
  "__SECRET_INTERNALS_DO_NOT_USE",
  "Adding different instances of a keyed plugin",
  "Duplicate use of selection JSON ID",
  "Yjs was already imported",
  " to a Fragment",
];

export class HostPeerPolicyError extends TaggedError("HostPeerPolicyError")<{
  message: string;
  packageDirectory: string;
  violations: readonly string[];
}> {
  constructor(packageDirectory: string, violations: readonly string[]) {
    super({
      message: `Host peer policy failed in ${packageDirectory}:\n${violations.join("\n")}`,
      packageDirectory,
      violations,
    });
  }
}

const walk = (directory: string): string[] => {
  const files: string[] = [];
  for (const entry of readdirSync(directory, { withFileTypes: true })) {
    const filename = path.join(directory, entry.name);
    if (entry.isDirectory()) files.push(...walk(filename));
    else if (entry.isFile()) files.push(filename);
  }
  return files;
};

export const discoverPublishedPackages = (
  repositoryDirectory: string,
): string[] =>
  readdirSync(path.join(repositoryDirectory, "packages"), {
    withFileTypes: true,
  })
    .filter((entry) => entry.isDirectory())
    .map((entry) => path.join(repositoryDirectory, "packages", entry.name))
    .filter((directory) => {
      const manifest = JSON.parse(
        readFileSync(path.join(directory, "package.json"), "utf8"),
      );
      return manifest.private !== true;
    })
    .sort();

const quotedSpecifier = (text: string): string | undefined =>
  /^(["'])([^"']*)\1$/u.exec(text)?.at(2);

/**
 * Singleton module specifiers a source names. `runtime` keeps only the ones
 * that survive type erasure: type-only imports, exports and `import()` types
 * are dropped.
 */
const singletonImports = (
  filename: string,
  source: string,
  { runtime }: { runtime: boolean },
): string[] => {
  const parsed = parseSourceModule(
    filename.endsWith(".vue") ? `${filename}.ts` : filename,
    source,
  );
  const { module } = parsed;
  const specifiers = [
    ...module.staticImports
      .filter(
        ({ entries }) =>
          !runtime ||
          entries.length === 0 ||
          entries.some(({ isType }) => !isType),
      )
      .map(({ moduleRequest }) => moduleRequest.value),
    ...module.staticExports.flatMap(({ entries }) =>
      entries.flatMap(({ moduleRequest, isType }) =>
        moduleRequest && !(runtime && isType) ? [moduleRequest.value] : [],
      ),
    ),
    ...module.dynamicImports.flatMap(({ moduleRequest }) => {
      const specifier = quotedSpecifier(
        source.slice(moduleRequest.start, moduleRequest.end),
      );
      return specifier === undefined ? [] : [specifier];
    }),
  ];
  walkAst(parsed.program, (node) => {
    if (
      !runtime &&
      node.type === "TSImportType" &&
      typeof node.source.value === "string"
    ) {
      specifiers.push(node.source.value);
    }
    // CommonJS `require("…")`, as a module scan of JavaScript reads it.
    const [argument] = node.type === "CallExpression" ? node.arguments : [];
    if (
      node.type === "CallExpression" &&
      node.callee.type === "Identifier" &&
      node.callee.name === "require" &&
      node.arguments.length === 1 &&
      argument?.type === "Literal" &&
      typeof argument.value === "string"
    ) {
      specifiers.push(argument.value);
    }
  });
  return specifiers.filter((specifier) => SINGLETON.test(specifier));
};

const sourceFiles = (packageDirectory: string): string[] =>
  walk(path.join(packageDirectory, "src"))
    .filter((filename) => /\.(?:ts|tsx|js|jsx|vue)$/u.test(filename))
    .filter(
      (filename) =>
        !TEST_SOURCE.test(path.relative(packageDirectory, filename)),
    );

const scriptSource = (filename: string): string => {
  const source = readFileSync(filename, "utf8");
  if (!filename.endsWith(".vue")) return source;
  return [...source.matchAll(/<script\b[^>]*>([\s\S]*?)<\/script>/gu)]
    .map((match) => match.at(1) ?? "")
    .join("\n");
};

export const assertHostPeerManifest = (packageDirectory: string): void => {
  const manifest = JSON.parse(
    readFileSync(path.join(packageDirectory, "package.json"), "utf8"),
  );
  const violations: string[] = [];
  for (const name of Object.keys(manifest.peerDependencies ?? {})) {
    if (
      SINGLETON.test(name) &&
      manifest.peerDependenciesMeta?.[name]?.optional === true
    ) {
      violations.push(`${name} must be a required peer`);
    }
  }
  for (const kind of ["dependencies", "optionalDependencies"]) {
    for (const name of Object.keys(manifest[kind] ?? {})) {
      if (SINGLETON.test(name))
        violations.push(`${name} must be a required peer, not ${kind}`);
    }
  }
  for (const filename of sourceFiles(packageDirectory)) {
    for (const specifier of singletonImports(filename, scriptSource(filename), {
      runtime: false,
    })) {
      const name = specifier.split("/").at(0);
      if (!name || !manifest.peerDependencies?.[name]) {
        violations.push(
          `${path.relative(packageDirectory, filename)} imports undeclared peer ${specifier}`,
        );
      } else if (manifest.peerDependenciesMeta?.[name]?.optional === true) {
        violations.push(`${name} must be a required peer`);
      }
    }
  }
  if (violations.length)
    throw new HostPeerPolicyError(packageDirectory, violations);
};

export const assertHostPeersExternal = (packageDirectory: string): void => {
  assertHostPeerManifest(packageDirectory);
  const expected = new Set<string>();
  for (const filename of sourceFiles(packageDirectory)) {
    if (filename.endsWith(".d.ts")) continue;
    for (const specifier of singletonImports(filename, scriptSource(filename), {
      runtime: true,
    })) {
      if (!specifier.endsWith(".css")) expected.add(specifier);
    }
  }
  const observed = new Set<string>();
  const violations: string[] = [];
  for (const filename of walk(path.join(packageDirectory, "dist"))) {
    if (!/\.(?:js|mjs|cjs)(?:\.map)?$/u.test(filename)) continue;
    const source = readFileSync(filename, "utf8");
    if (filename.endsWith(".map")) {
      const map = JSON.parse(source);
      for (const origin of map.sources ?? []) {
        if (
          /node_modules\/(?:prosemirror-[a-z-]+|yjs|y-prosemirror|y-protocols|react|react-dom)(?:\/|$)/u.test(
            origin,
          )
        ) {
          violations.push(
            `${filename} includes bundled singleton source ${origin}`,
          );
        }
      }
      continue;
    }
    for (const specifier of singletonImports(filename, source, {
      runtime: false,
    }))
      observed.add(specifier);
    for (const sentinel of BUNDLED_SENTINELS) {
      if (source.includes(sentinel))
        violations.push(
          `${filename} includes bundled singleton code (${sentinel})`,
        );
    }
  }
  for (const specifier of expected) {
    if (!observed.has(specifier))
      violations.push(
        `Runtime import ${specifier} is missing from built externals`,
      );
  }
  if (violations.length)
    throw new HostPeerPolicyError(packageDirectory, violations);
};

export const assertHostPeerResolution = (consumerDirectory: string): void => {
  const hostRequire = createRequire(
    path.join(consumerDirectory, "package.json"),
  );
  const scopeDirectory = path.join(consumerDirectory, "node_modules", "@stll");
  const violations: string[] = [];
  let inspected = 0;
  for (const entry of readdirSync(scopeDirectory)) {
    const manifestFile = path.join(scopeDirectory, entry, "package.json");
    if (!existsSync(manifestFile)) continue;
    const manifest = JSON.parse(readFileSync(manifestFile, "utf8"));
    if (manifest.private === true) continue;
    const packageRequire = createRequire(realpathSync(manifestFile));
    for (const name of Object.keys(manifest.peerDependencies ?? {})) {
      if (!SINGLETON.test(name)) continue;
      inspected++;
      const specifier = name === "y-protocols" ? "y-protocols/awareness" : name;
      const hostPath = realpathSync(hostRequire.resolve(specifier));
      const packagePath = realpathSync(packageRequire.resolve(specifier));
      if (hostPath !== packagePath) {
        violations.push(
          `${manifest.name} resolves ${specifier} to ${packagePath}; host uses ${hostPath}`,
        );
      }
    }
  }
  if (!inspected)
    violations.push("No installed singleton peers were inspected");
  if (violations.length)
    throw new HostPeerPolicyError(consumerDirectory, violations);
};
