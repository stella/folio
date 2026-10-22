import { describe, expect, test } from "bun:test";

import { resolvedWasmBuildGraphDescription } from "./lib/rust-wasm-artifact.ts";

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
