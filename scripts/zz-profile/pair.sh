#!/bin/bash
# Scratch only: CPU profile pair of the canonical per-key path, base vs patch, alternating rounds on one runner.
set -u
HERE=$(cd "$(dirname "$0")" && pwd)
PATCH=$(cd "$HERE/../.." && pwd)
BASE=$1; OUT=$2; ROUNDS=${3:-3}
NAMES="stage normalizeListModelRendering normalizeBlock compileEditorIntent toProseDoc project prepareReplace"
mkdir -p "$OUT"
for round in $(seq 1 "$ROUNDS"); do
  for doc in list-3500 contract-100; do
    for side in base patch; do
      if [ $side = base ]; then wt=$BASE; else wt=$PATCH; fi
      dir=$OUT/$side-$doc-$round; mkdir -p "$dir"
      cp "$HERE/zz-perkey.run.txt" "$wt/packages/core/src/controller/zz-perkey.run.ts"
      (cd "$wt/packages/core" && ZZ_DOC=$doc ZZ_KEYS=120 bun --cpu-prof --cpu-prof-dir="$dir" src/controller/zz-perkey.run.ts) > "$dir/stdout.txt" 2>&1
      rm -f "$wt/packages/core/src/controller/zz-perkey.run.ts"
      prof=$(ls "$dir"/*.cpuprofile | head -1)
      echo "== $side $doc round $round"
      grep '^ZZ' "$dir/stdout.txt"
      python3 -I "$HERE/incl.py" "$prof" 120 $NAMES
      python3 -I "$HERE/top.py" "$prof" 120 stage,normalizeListModelRendering 12 | head -14
    done
  done
done
