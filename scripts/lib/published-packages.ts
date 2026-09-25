// The packages folio publishes, in the order the root `build` script walks
// them (a later package's build may read an earlier one's dist).
//
// Shared by every gate that inspects published artifacts, so the set cannot
// drift from the release set by being hand-listed in each gate.

import path from "node:path";

const repoRoot = path.resolve(import.meta.dir, "..", "..");

export type PublishedPackage = {
  /** Its `packages/<slug>` directory name; also how a gate's CLI names it. */
  slug: string;
  name: string;
  root: string;
};

const published = (slug: string, name: string): PublishedPackage => ({
  slug,
  name,
  root: path.join(repoRoot, "packages", slug),
});

export const PUBLISHED_PACKAGES: readonly PublishedPackage[] = [
  published("docx-core", "@stll/docx-core"),
  published("core", "@stll/folio-core"),
  published("react", "@stll/folio-react"),
  published("agents", "@stll/folio-agents"),
  published("cli", "@stll/folio-cli"),
  published("vue", "@stll/folio-vue"),
  published("nuxt", "@stll/folio-nuxt"),
];
