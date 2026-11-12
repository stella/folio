#!/bin/bash
# Scratch only: canonical vs legacy, alternating engines (order swapped each round) on one runner.
# usage: rounds.sh <out dir> <rounds>
set -u
OUT=$1; ROUNDS=$2
mkdir -p "$OUT"
for round in $(seq 1 "$ROUNDS"); do
  if [ $((round % 2)) = 1 ]; then order=legacy,canonical; else order=canonical,legacy; fi
  FOLIO_ZZ_ROUND=$round FOLIO_ZZ_ORDER=$order FOLIO_ZZ_OUT="$OUT/results.jsonl" \
    bunx playwright test -c zz-flip.config.ts > "$OUT/round-$round.log" 2>&1
  echo "== round $round ($order) rc=$?"
  grep '^ZZFLIP\|failed\|Error' "$OUT/round-$round.log" | cut -c1-400
done
# One extra pass with CPU profiles of the measured keys (timings from this pass are not used).
FOLIO_ZZ_ROUND=profile FOLIO_ZZ_PROFILE="$OUT/profiles" FOLIO_ZZ_OUT="$OUT/profile-pass.jsonl" \
  bunx playwright test -c zz-flip.config.ts > "$OUT/profile.log" 2>&1
echo "== profile pass rc=$?"
grep '^ZZFLIP' "$OUT/profile.log" | cut -c1-400
