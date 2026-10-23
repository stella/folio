// Derive publication membership from workspace manifests and build dependencies first.
import { readFileSync } from "node:fs";
import path from "node:path";

import { publishedPackagePaths } from "../ci-api-plan";

const repoRoot = path.resolve(import.meta.dir, "..", "..");

export type PublishedPackage = {
  slug: string;
  name: string;
  root: string;
};

export const discoverPublishedPackages = (
  root = repoRoot,
): PublishedPackage[] => {
  const packages = publishedPackagePaths(root).map((directory) => {
    const manifest = JSON.parse(
      readFileSync(path.join(root, directory, "package.json"), "utf8"),
    );
    return {
      slug: path.basename(directory),
      name: manifest.name,
      root: path.join(root, directory),
      dependencies: Object.keys({
        ...manifest.dependencies,
        ...manifest.peerDependencies,
      }),
    };
  });
  const ordered: PublishedPackage[] = [];
  const pending = new Map(packages.map((pkg) => [pkg.name, pkg]));
  while (pending.size > 0) {
    const ready = [...pending.values()].filter((pkg) =>
      pkg.dependencies.every((name) => !pending.has(name)),
    );
    if (ready.length === 0)
      throw new TypeError("Published workspace dependencies contain a cycle");
    for (const { slug, name, root: packageRoot } of ready) {
      ordered.push({ slug, name, root: packageRoot });
      pending.delete(name);
    }
  }
  return ordered;
};

export const PUBLISHED_PACKAGES: readonly PublishedPackage[] =
  discoverPublishedPackages();
