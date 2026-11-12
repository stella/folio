#!/usr/bin/env bash
# Fails when a change adds a changeset that a release already consumed.
#
# A release ("chore: version packages") deletes every changeset it publishes.
# A branch that added a changeset before that release and merges the released
# main afterwards keeps its own copy, so merging it would publish the same note
# again in the next release. This compares the changesets the change adds with
# those every release on the main line deleted.
#
# usage: check-released-changesets.sh <main ref> <base ref> [<head ref>]
#   main ref  the line whose release commits count (origin/main)
#   base ref  the branch the change merges into
#   head ref  the change (default HEAD)
set -euo pipefail

main_ref=$1
base_ref=$2
head_ref=${3:-HEAD}

released=$(git log "$main_ref" --diff-filter=D --name-only --format= \
  --grep='^chore: version packages' -- '.changeset/*.md' | sort -u)
merge_base=$(git merge-base "$base_ref" "$head_ref")
added=$(git diff --name-only --diff-filter=A "$merge_base" "$head_ref" \
  -- '.changeset/*.md' | sort -u)

again=$(comm -12 <(printf '%s\n' "$released") <(printf '%s\n' "$added") |
  sed '/^$/d')
if [ -n "$again" ]; then
  while IFS= read -r file; do
    echo "::error file=$file::A release already published and deleted $file; remove it (a merge of the released main brought it back) or give a new note its own name."
  done <<<"$again"
  exit 1
fi
echo "No changeset a release already published is added."
