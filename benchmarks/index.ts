/**
 * Benchmark runner.
 *
 * `bun run bench` executes every group under Bun and prints a table. In CI the
 * CodSpeed action runs the same file under Node; the `withCodSpeed`-wrapped
 * benches then report to the instrumentation instead of printing locally.
 *
 * Bench files are runtime-agnostic (fixtures read via `node:fs`), so the same
 * code runs under both Bun and Node.
 */
import type { Bench } from "tinybench";

import { editorTransactionBench } from "./editor-transaction.bench";
import { lineBreakBench } from "./line-break.bench";
import { markdownBench } from "./markdown.bench";
import { parseBench } from "./parse.bench";
import { proseMirrorBench } from "./prosemirror.bench";
import { toBenchmarkRow } from "./report";
import { serializeBench } from "./serialize.bench";

type Group = {
  readonly name: string;
  readonly make: () => Bench | Promise<Bench>;
};

const GROUPS: readonly Group[] = [
  { name: "parse · DOCX → model (folio)", make: parseBench },
  { name: "serialize · model → DOCX (folio)", make: serializeBench },
  { name: "markdown · model ↔ Markdown (folio)", make: markdownBench },
  { name: "prosemirror · model ↔ ProseMirror (folio)", make: proseMirrorBench },
  {
    name: "line breaking · paragraph text → wrap offsets (folio)",
    make: lineBreakBench,
  },
  {
    name: "editor · one many-step transaction → plugin stack (folio)",
    make: editorTransactionBench,
  },
];

for (const group of GROUPS) {
  const bench = await group.make();
  await bench.run();
  console.log(`\n=== ${group.name} ===`);
  console.table(bench.tasks.map(toBenchmarkRow));
}
