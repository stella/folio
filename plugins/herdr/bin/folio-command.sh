#!/usr/bin/env bash
# The folio command this plugin runs: $FOLIO_BIN when set (for example a
# checkout: "bun /path/to/folio/packages/cli/src/bin.ts"), `folio` on PATH,
# or the @stll/folio-cli version this plugin was released with, through npx.
# Sourced by the plugin's scripts; sets the FOLIO array.
# shellcheck disable=SC2034 # FOLIO is read by the scripts that source this file.

if [[ -n "${FOLIO_BIN:-}" ]]; then
  read -r -a FOLIO <<<"$FOLIO_BIN"
elif command -v folio >/dev/null 2>&1; then
  FOLIO=(folio)
else
  # `herdr plugin install` checks out the whole folio repo and runs this
  # plugin's scripts with a cwd inside it, so a plain npx here would resolve
  # the monorepo's own unbuilt @stll/folio-cli workspace package instead of
  # fetching the published one. Run it from outside any checkout instead.
  FOLIO=(bash -c 'cd "$1" && shift && exec npx -y @stll/folio-cli@0.3.0 "$@"' bash "${HOME:-/tmp}")
fi
