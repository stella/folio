#!/bin/bash
# Scratch only: main against the branch, canonical engine, alternating (order swapped each round) on one runner.
# usage: rounds.sh <out dir> <rounds>   (main served on 4200, branch on 4201)
set -u
OUT=$1; ROUNDS=$2
mkdir -p "$OUT"
for round in $(seq 1 "$ROUNDS"); do
  if [ $((round % 2)) = 1 ]; then order="main branch"; else order="branch main"; fi
  for variant in $order; do
    if [ "$variant" = main ]; then port=4200; else port=4201; fi
    FOLIO_ZZ_ROUND=$round FOLIO_ZZ_VARIANT=$variant FOLIO_ZZ_ORDER=canonical \
      FOLIO_ZZ_BASE="http://localhost:$port" FOLIO_ZZ_OUT="$OUT/results.jsonl" \
      bunx playwright test -c zz-flip.config.ts > "$OUT/round-$round-$variant.log" 2>&1
    echo "== round $round $variant rc=$?"
    grep '^ZZFLIP\|failed\|Error' "$OUT/round-$round-$variant.log" | cut -c1-300
  done
done
