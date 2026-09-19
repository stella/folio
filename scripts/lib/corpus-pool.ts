/**
 * A bounded pool of corpus workers with a per-file deadline.
 *
 * Each worker holds one file at a time. If it does not answer within the
 * deadline, or exits while holding a file, the pool kills it, records what
 * happened and starts a replacement: neither is a reason to lose the rest of
 * the run.
 */

import path from "node:path";

import { REPOSITORY_ROOT } from "./corpus-manifest";
import type { CorpusCheckResult } from "./corpus-check";
import { EXTENDED_CORPUS_INVARIANTS } from "./corpus-invariants/contract";
import { CORPUS_INVARIANTS, type CorpusFailure, failureFromAssertion } from "./corpus-signature";

const WORKER_ENTRY = path.join(REPOSITORY_ROOT, "scripts", "lib", "corpus-worker.ts");

/** Advisory per-file time budgets, forwarded to the extended invariants. */
export type CorpusBudgets = {
  invariantBudgetMs: number;
  fileBudgetMs: number;
};

export type CorpusTask = {
  sourceId: string;
  relativePath: string;
  sha256: string;
  absolutePath: string;
};

/**
 * How a file's run ended.
 *
 * `aborted` and `watchdog-expired` are both "the worker stopped answering",
 * and they mean opposite things: a worker that exited died on the file, which
 * is a fact about the file and gates, while a deadline that expired cannot
 * tell a hung worker from a slow machine. The second is a truncation by
 * another name, so it is kept apart here rather than reconstructed from the
 * invariant its failure happens to carry.
 */
export type CorpusTaskOutcome =
  | CorpusCheckResult
  | { kind: "aborted"; failures: CorpusFailure[] }
  | { kind: "watchdog-expired"; failures: CorpusFailure[] };

type PendingLine = { line: string } | { line: null; ended: "eof" | "timeout" };

const readLines = async function* (
  stream: ReadableStream<Uint8Array>,
): AsyncGenerator<string, void, undefined> {
  const decoder = new TextDecoder();
  let buffer = "";
  for await (const chunk of stream) {
    buffer += decoder.decode(chunk, { stream: true });
    let newline = buffer.indexOf("\n");
    while (newline !== -1) {
      yield buffer.slice(0, newline);
      buffer = buffer.slice(newline + 1);
      newline = buffer.indexOf("\n");
    }
  }
};

const spawnWorker = () =>
  Bun.spawn(["bun", WORKER_ENTRY], {
    cwd: REPOSITORY_ROOT,
    stdin: "pipe",
    stdout: "pipe",
    stderr: "ignore",
  });

class PooledWorker {
  #child: ReturnType<typeof spawnWorker> | null = null;
  #lines: AsyncGenerator<string, void, undefined> | null = null;
  #nextId = 0;

  #start(): void {
    const child = spawnWorker();
    this.#child = child;
    this.#lines = readLines(child.stdout);
  }

  kill(): void {
    this.#child?.kill();
    this.#child = null;
    this.#lines = null;
  }

  async run(
    task: CorpusTask,
    timeoutMs: number,
    budgets: CorpusBudgets,
  ): Promise<CorpusTaskOutcome> {
    if (this.#child === null || this.#lines === null) {
      this.#start();
    }
    const child = this.#child;
    const lines = this.#lines;
    if (child === null || lines === null) {
      return abortedOutcome("the worker could not be started");
    }

    this.#nextId += 1;
    const id = this.#nextId;
    child.stdin.write(`${JSON.stringify({ id, path: task.absolutePath, ...budgets })}\n`);
    child.stdin.flush();

    const timeout = Promise.withResolvers<PendingLine>();
    const timer = setTimeout(() => {
      timeout.resolve({ line: null, ended: "timeout" });
    }, timeoutMs);
    const answered = await Promise.race([
      lines
        .next()
        .then(
          (result): PendingLine =>
            result.done === true ? { line: null, ended: "eof" } : { line: result.value },
        ),
      timeout.promise,
    ]);
    clearTimeout(timer);

    if (answered.line === null) {
      this.kill();
      if (answered.ended === "eof") {
        return abortedOutcome("the worker exited without a verdict");
      }
      return {
        kind: "watchdog-expired",
        failures: [
          failureFromAssertion(
            EXTENDED_CORPUS_INVARIANTS.performance,
            `no verdict within ${Math.round(timeoutMs / 1000)}s (watchdog expired)`,
          ),
        ],
      };
    }
    const parsed = JSON.parse(answered.line) as { id: number; result: CorpusCheckResult };
    if (parsed.id !== id) {
      this.kill();
      return abortedOutcome("the worker answered out of order");
    }
    return parsed.result;
  }
}

const abortedOutcome = (detail: string): CorpusTaskOutcome => ({
  kind: "aborted",
  failures: [failureFromAssertion(CORPUS_INVARIANTS.completes, detail)],
});

export type RunPoolOptions = {
  tasks: readonly CorpusTask[];
  concurrency: number;
  timeoutMs: number;
  budgets: CorpusBudgets;
  onOutcome: (task: CorpusTask, outcome: CorpusTaskOutcome) => void;
};

export const runCorpusPool = async ({
  tasks,
  concurrency,
  timeoutMs,
  budgets,
  onOutcome,
}: RunPoolOptions): Promise<void> => {
  let next = 0;
  const worker = async (): Promise<void> => {
    const pooled = new PooledWorker();
    while (next < tasks.length) {
      const task = tasks[next];
      next += 1;
      if (task === undefined) {
        break;
      }
      // oxlint-disable-next-line no-await-in-loop -- a worker holds one file at a time by design
      onOutcome(task, await pooled.run(task, timeoutMs, budgets));
    }
    pooled.kill();
  };
  await Promise.all(Array.from({ length: Math.min(concurrency, tasks.length) }, worker));
};
