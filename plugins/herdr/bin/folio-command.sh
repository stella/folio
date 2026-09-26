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
  FOLIO=(npx -y @stll/folio-cli@0.0.0)
fi
