import { describe, expect, test } from "bun:test";

import {
  manifestBuildDescription,
  resolvedWasmBuildGraphDescription,
} from "./lib/rust-wasm-artifact.ts";

type DependencyKind = "dev" | "build" | null;

const packageOf = (name: string, version = "1.0.0") => ({
  id: `${name} ${version}`,
  name,
  version,
  source: name === "app" || name === "other" ? null : "registry+https://example.invalid/index",
  checksum: name === "app" || name === "other" ? null : `${name}-${version}-checksum`,
});

const dependencyOn = (name: string, version: string, kind: DependencyKind) => ({
  pkg: `${name} ${version}`,
  dep_kinds: [{ kind }],
});

const fixture = ({
  dependencyVersion = "1.0.0",
  transitiveVersion = "1.0.0",
  buildVersion = "1.0.0",
  devVersion = "1.0.0",
  otherVersion = "1.0.0",
} = {}) => {
  const packages = [
    packageOf("app"),
    packageOf("dependency", dependencyVersion),
    packageOf("transitive", transitiveVersion),
    packageOf("build-helper", buildVersion),
    packageOf("dev-helper", devVersion),
    packageOf("other"),
    packageOf("other-dependency", otherVersion),
  ];
  return {
    packages,
    resolve: {
      nodes: [
        {
          id: "app 1.0.0",
          deps: [
            dependencyOn("dependency", dependencyVersion, null),
            dependencyOn("build-helper", buildVersion, "build"),
            dependencyOn("dev-helper", devVersion, "dev"),
          ],
        },
        {
          id: `dependency ${dependencyVersion}`,
          deps: [dependencyOn("transitive", transitiveVersion, null)],
        },
        { id: `transitive ${transitiveVersion}`, deps: [] },
        { id: `build-helper ${buildVersion}`, deps: [] },
        { id: `dev-helper ${devVersion}`, deps: [] },
        {
          id: "other 1.0.0",
          deps: [dependencyOn("other-dependency", otherVersion, null)],
        },
        { id: `other-dependency ${otherVersion}`, deps: [] },
      ],
    },
  };
};

const describeFixture = (metadata = fixture()) =>
  resolvedWasmBuildGraphDescription(metadata, "app");

describe("resolvedWasmBuildGraphDescription", () => {
  test("ignores added and bumped dev-dependencies", () => {
    const baseline = fixture();
    const bumped = fixture({ devVersion: "2.0.0" });
    bumped.packages.push(packageOf("another-dev-helper"));
    bumped.resolve.nodes.at(0)?.deps.push(dependencyOn("another-dev-helper", "1.0.0", "dev"));
    bumped.resolve.nodes.push({ id: "another-dev-helper 1.0.0", deps: [] });

    expect(describeFixture(bumped)).toBe(describeFixture(baseline));
  });

  test("changes for direct and transitive normal dependency bumps", () => {
    const baseline = describeFixture();

    expect(describeFixture(fixture({ dependencyVersion: "2.0.0" }))).not.toBe(baseline);
    expect(describeFixture(fixture({ transitiveVersion: "2.0.0" }))).not.toBe(baseline);
  });

  test("changes for build-dependency bumps", () => {
    expect(describeFixture(fixture({ buildVersion: "2.0.0" }))).not.toBe(describeFixture());
  });

  test("changes for a host-only dependency of a build dependency", () => {
    const withHostDependency = (version: string) => {
      const metadata = fixture();
      metadata.packages.push(packageOf("host-only", version));
      const buildHelper = metadata.resolve.nodes.find(({ id }) => id === "build-helper 1.0.0");
      buildHelper?.deps.push({
        pkg: `host-only ${version}`,
        dep_kinds: [{ kind: null, target: "cfg(unix)" }],
      });
      metadata.resolve.nodes.push({ id: `host-only ${version}`, deps: [] });
      return metadata;
    };

    expect(describeFixture(withHostDependency("2.0.0"))).not.toBe(
      describeFixture(withHostDependency("1.0.0")),
    );
  });

  test("ignores dependencies owned by another workspace member", () => {
    expect(describeFixture(fixture({ otherVersion: "2.0.0" }))).toBe(describeFixture());
  });

  test("is independent of package and node order", () => {
    const reordered = fixture();
    reordered.packages.reverse();
    reordered.resolve.nodes.reverse();
    for (const node of reordered.resolve.nodes) {
      node.deps.reverse();
    }

    expect(describeFixture(reordered)).toBe(describeFixture());
  });
});

const manifest = ({
  dependency = '"1"',
  devDependency = '"8"',
  feature = '["dep:js-sys"]',
  targetDevDependency = '"1"',
} = {}) => `[package]
name = "app"

[features]
wasm = ${feature}

[dependencies]
serde = ${dependency}

[dev-dependencies]
zip = ${devDependency}

[target.'cfg(unix)'.dependencies]
libc = "0.2"

[target.'cfg(unix)'.dev-dependencies]
tempfile = ${targetDevDependency}
`;

describe("manifestBuildDescription", () => {
  test("ignores dev-dependency edits, including target-specific ones", () => {
    const baseline = manifestBuildDescription(manifest());

    expect(manifestBuildDescription(manifest({ devDependency: '{ version = "9.0.1" }' }))).toBe(
      baseline,
    );
    expect(manifestBuildDescription(manifest({ targetDevDependency: '"3"' }))).toBe(baseline);
  });

  test("changes for dependency and feature edits", () => {
    const baseline = manifestBuildDescription(manifest());

    expect(manifestBuildDescription(manifest({ dependency: '"2"' }))).not.toBe(baseline);
    expect(
      manifestBuildDescription(manifest({ feature: '["dep:js-sys", "dep:web-sys"]' })),
    ).not.toBe(baseline);
  });

  test("is independent of table and key order", () => {
    const reordered = `[dependencies]
serde = "1"

[target.'cfg(unix)'.dev-dependencies]
tempfile = "1"

[target.'cfg(unix)'.dependencies]
libc = "0.2"

[features]
wasm = ["dep:js-sys"]

[package]
name = "app"
`;

    expect(manifestBuildDescription(reordered)).toBe(manifestBuildDescription(manifest()));
  });
});
