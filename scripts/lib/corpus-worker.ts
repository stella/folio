/**
 * One corpus file per line in, one verdict per line out.
 *
 * The invariants run in a child process because the failures the gate looks for
 * include the ones a library cannot report: an unbounded loop, a stack that
 * overflows, an allocation that aborts the runtime. The orchestrator kills a
 * worker that stops answering and records the file, which an in-process runner
 * could not do.
 */

import { Result } from "better-result";

import { runCorpusChecks } from "./corpus-check";
import { CORPUS_INVARIANTS, failureFromError } from "./corpus-signature";

type WorkerRequest = {
  id: number;
  path: string;
  invariantBudgetMs: number;
  fileBudgetMs: number;
};

for await (const line of console) {
  if (line.length === 0) {
    continue;
  }
  const request = JSON.parse(line) as WorkerRequest;
  const bytes = await Result.tryPromise({
    try: async () => new Uint8Array(await Bun.file(request.path).arrayBuffer()),
    catch: (cause: unknown) => cause,
  });
  const result = bytes.isErr()
    ? {
        kind: "checked" as const,
        failures: [failureFromError(CORPUS_INVARIANTS.completes, bytes.error)],
        producer: "unknown",
        cost: { bytes: 0, parseMs: 0, peakRssBytes: process.memoryUsage.rss() },
        timings: {},
      }
    : await runCorpusChecks(bytes.value, {
        invariantBudgetMs: request.invariantBudgetMs,
        fileBudgetMs: request.fileBudgetMs,
      });
  process.stdout.write(`${JSON.stringify({ id: request.id, result })}\n`);
}
