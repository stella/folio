#!/usr/bin/env bash
# Link handler for stll.folio: turn the Ctrl-clicked link into a local .docx
# path and open the plugin's preview pane on it, split to the right of the
# clicked pane without taking focus.
set -euo pipefail

link="${HERDR_PLUGIN_CLICKED_URL:-${1:-}}"
if [[ -z "$link" ]]; then
  echo "no link given (HERDR_PLUGIN_CLICKED_URL is empty)" >&2
  exit 1
fi

# file:///a/b%20c.docx and file://localhost/a/b.docx name /a/b c.docx and /a/b.docx.
path="$link"
if [[ "$path" =~ ^[Ff][Ii][Ll][Ee]:// ]]; then
  path="${path:7}"
  if [[ "$path" != /* ]]; then
    host="${path%%/*}"
    # `ls --hyperlink` names this machine by its host name.
    if [[ "$host" != "localhost" && "$host" != "$(hostname)" && "$host" != "$(hostname -s)" ]]; then
      echo "not a local file: $link" >&2
      exit 1
    fi
    path="/${path#*/}"
  fi
  path="$(printf '%b' "${path//%/\\x}")"
fi
if [[ ! "$path" =~ \.[Dd][Oo][Cc][Xx]$ ]]; then
  echo "not a .docx: $link" >&2
  exit 1
fi

# A bare (non-file://) link can be relative to the clicked pane, but this
# action itself runs with the plugin directory as its cwd, and the preview
# pane's `folio` command may run through npx from yet another cwd. Make the
# path absolute now, against the clicked pane's cwd when herdr reports it.
if [[ "$path" != /* ]]; then
  cwd="$(printf '%s' "${HERDR_PLUGIN_CONTEXT_JSON:-}" | grep -o '"focused_pane_cwd":"[^"]*"' | head -n 1)" || true
  cwd="${cwd#\"focused_pane_cwd\":\"}"
  cwd="${cwd%\"}"
  if [[ -z "$cwd" ]]; then
    cwd="$PWD"
    echo "warning: clicked pane's cwd is unavailable; resolving \"$path\" against $cwd" >&2
  fi
  path="$cwd/$path"
fi

args=(plugin pane open --plugin "${HERDR_PLUGIN_ID:-stll.folio}" --entrypoint preview
  --placement split --direction right --no-focus --env "FOLIO_DOCX=$path")
[[ -n "${HERDR_PANE_ID:-}" ]] && args+=(--target-pane "$HERDR_PANE_ID")

"${HERDR_BIN_PATH:-herdr}" "${args[@]}"
