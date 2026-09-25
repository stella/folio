#!/usr/bin/env bash
# Preview pane for stll.folio. With terminal-browser installed, run
# `folio serve` on the document and open its live preview; the page follows
# the file as it changes. Without it, render one page with `folio render` and
# show it through the kitty graphics protocol. Neither path writes the document.
set -euo pipefail

doc="${FOLIO_DOCX:-${1:-}}"
if [[ -z "$doc" ]]; then
  echo "no document given (FOLIO_DOCX is empty)" >&2
  exit 1
fi
# shellcheck source=folio-command.sh
source "$(dirname "${BASH_SOURCE[0]}")/folio-command.sh"

if command -v terminal-browser >/dev/null 2>&1; then
  log="$(mktemp)"
  "${FOLIO[@]}" serve "$doc" --output json >"$log" 2>&1 &
  server=$!
  trap 'kill "$server" 2>/dev/null; rm -f "$log"' EXIT
  url=""
  for _ in $(seq 1 300); do
    url="$(grep -o '"url":"[^"]*"' "$log" | head -n 1 | sed -e 's/^"url":"//' -e 's/"$//')" || true
    [[ -n "$url" ]] && break
    if ! kill -0 "$server" 2>/dev/null; then
      cat "$log" >&2
      exit 1
    fi
    sleep 0.1
  done
  if [[ -z "$url" ]]; then
    echo "folio serve did not start" >&2
    exit 1
  fi
  terminal-browser open "$url" --no-merge
  exit 0
fi

# Kitty graphics fallback: one page at a time. n/p change page, r re-renders,
# q quits. `kitten icat` sizes the image to the pane when it is available.
show_png() {
  if command -v kitten >/dev/null 2>&1; then
    kitten icat --clear --transfer-mode=stream "$1"
    return
  fi
  local data chunk more first=1
  data="$(base64 <"$1" | tr -d '\n')"
  while [[ -n "$data" ]]; do
    chunk="${data:0:4096}"
    data="${data:4096}"
    more=1
    [[ -z "$data" ]] && more=0
    if ((first)); then
      printf '\033_Gf=100,a=T,m=%d;%s\033\\' "$more" "$chunk"
      first=0
    else
      printf '\033_Gm=%d;%s\033\\' "$more" "$chunk"
    fi
  done
  printf '\n'
}

workdir="$(mktemp -d)"
trap 'rm -rf "$workdir"' EXIT
page=1
while true; do
  printf '\033[2J\033[H'
  png="$workdir/page-$page.png"
  rm -f "$png"
  if result="$("${FOLIO[@]}" render "$doc" -o "$png" --page "$page" --output json 2>&1)"; then
    show_png "$png"
  else
    printf '%s\n' "$result"
    ((page > 1)) && page=$((page - 1))
  fi
  printf 'page %d  [n]ext [p]revious [r]eload [q]uit\n' "$page"
  read -r -n 1 -s key || exit 0
  case "$key" in
    n) page=$((page + 1)) ;;
    p) ((page > 1)) && page=$((page - 1)) ;;
    q) exit 0 ;;
    *) ;;
  esac
done
