#!/bin/bash
# Scratch only: getDocument and save timings, base vs patch, alternating rounds on one runner.
# usage: pair.sh <base tree> <out dir> <rounds>
set -u
HERE=$(cd "$(dirname "$0")" && pwd)
PATCH=$(cd "$HERE/../.." && pwd)
BASE=$1; OUT=$2; ROUNDS=$3
mkdir -p "$OUT"
for round in $(seq 1 "$ROUNDS"); do
  for doc in list-3500 contract-100; do
    for side in base patch; do
      if [ $side = base ]; then wt=$BASE; else wt=$PATCH; fi
      log=$OUT/$side-$doc-$round.txt
      cp "$HERE/zz-getdoc.test.txt" "$wt/packages/react/src/components/zz-getdoc.test.tsx"
      (cd "$wt/packages/react" && ZZ_DOC=$doc bun test src/components/zz-getdoc.test.tsx) > "$log" 2>&1
      rm -f "$wt/packages/react/src/components/zz-getdoc.test.tsx"
      echo "== $side $doc round $round"
      grep '^ZZ' "$log" || tail -30 "$log"
    done
  done
done
