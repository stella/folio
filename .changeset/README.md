# Changesets

This folder is managed by [Changesets](https://github.com/changesets/changesets).

Every PR that changes published source under `packages/{docx-core,core,react,
agents,vue,nuxt}/src` must ship a changeset describing the release:

```sh
bunx changeset
```

Pick the changed packages, a bump level (`patch` / `minor` / `major`), and write
a one-line summary — it becomes the changelog entry. Commit the generated
`.changeset/*.md` file with your PR.

If a source change genuinely needs no release (comments, internal refactors that
do not affect the published API), record that explicitly:

```sh
bunx changeset --empty
```

The shared Changesets policy check fails a PR that touches those source paths
without one of the above.

## How a release happens

1. PRs merge to `main`, each carrying its changeset(s).
2. `release-pr.yml` maintains a **"Version Packages"** PR that applies the
   pending changesets: it bumps the affected `package.json` versions, updates the
   changelogs, and re-syncs `bun.lock`.
3. Merging that PR lands the version bumps on `main`. `publish.yml` builds and
   packs all six public packages without a publishing credential, then delegates
   registry state, dependency ordering, OIDC publishing, and per-package GitHub
   Releases to the versioned shared workflow.

Changesets never publishes here; `publish.yml` remains the sole publish
mechanism.
