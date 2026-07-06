#!/usr/bin/env bun
// CI gate: a PR that edits the published source of @stll/folio-core,
// @stll/folio-react, @stll/folio-vue, or @stll/folio-nuxt must ship a changeset
// (`.changeset/*.md`) so the change gets a version bump + changelog entry, OR
// an explicit empty changeset (`bunx changeset --empty`) to record on purpose
// that no release is needed.
//
// Scoped to `packages/{core,react,vue,nuxt}/src` deliberately: it ignores
// tests, docs, and config, and it passes the bot's "Version Packages" PR — that
// PR consumes changesets and bumps package.json, but never touches src, so it
// needs none.
//
// Compares the branch against its merge base with the base branch (three-dot
// diff), so it only inspects what this PR introduces. Override the base ref
// with CHANGESET_BASE_REF when running locally against a different branch.

import { $ } from "bun";

const BASE = process.env.CHANGESET_BASE_REF ?? "origin/main";
const SRC_RE = /^packages\/(?:core|react|vue|nuxt)\/src\//;
// Any `.changeset/*.md` except the folder's own README is a changeset entry
// (an empty changeset produced by `bunx changeset --empty` counts too).
const CHANGESET_RE = /^\.changeset\/(?!README\.md$)[^/]+\.md$/;

// Best-effort: make sure the base ref is present locally. A full checkout
// already has it; ignore failure here (e.g. offline local runs) — if the ref
// really is missing, the `git diff` below fails loudly.
await $`git fetch --no-tags origin ${BASE.replace(/^origin\//, "")}`.nothrow().quiet();

const diff = async (filter: string): Promise<string[]> => {
  // `.nothrow()` + an explicit exit-code check: a failed `git diff` (missing
  // base ref, shallow clone, malformed BASE) must fail the job loudly. Letting
  // it read as an empty diff would silently bypass the whole gate.
  const result = await $`git diff --name-only --diff-filter=${filter} ${BASE}...HEAD`
    .nothrow()
    .quiet();
  if (result.exitCode !== 0) {
    console.error(
      `changeset check: \`git diff ${BASE}...HEAD\` failed (exit ${result.exitCode}); cannot verify changesets.`,
    );
    console.error(result.stderr.toString());
    process.exit(1);
  }
  return result.stdout
    .toString()
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean);
};

// Additions, modifications, renames, deletions to src all count as a
// release-relevant change (removing an export is a release too).
const srcChanged = (await diff("ACMRD")).some((file) => SRC_RE.test(file));
if (!srcChanged) {
  console.log("changeset check: no packages/{core,react,vue,nuxt}/src changes; skipping.");
  process.exit(0);
}

// The PR must ADD its own changeset; a changeset already merged to main does
// not cover a later PR.
const changesetAdded = (await diff("A")).some((file) => CHANGESET_RE.test(file));
if (changesetAdded) {
  console.log("changeset check: source change has a changeset. OK.");
  process.exit(0);
}

console.error(
  [
    "Missing changeset.",
    "",
    "This PR edits packages/{core,react,vue,nuxt}/src (published source)",
    "but adds no changeset. Add one so the change gets a version bump and a",
    "changelog entry:",
    "",
    "    bunx changeset",
    "",
    "If the change genuinely needs no release (comments, internal-only refactor),",
    "record that explicitly with an empty changeset:",
    "",
    "    bunx changeset --empty",
    "",
    "Then commit the generated .changeset/*.md file.",
  ].join("\n"),
);
process.exit(1);
