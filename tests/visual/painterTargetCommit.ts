import { Result } from "better-result";

type PainterTargetOptions<T> = {
  commit: () => void;
  read: () => T | null;
  subscribe: (listener: () => void) => () => void;
};

/** Commit current document paint before reading; numeric hit tests can match stale pages. */
export const resolvePainterTarget = <T>({
  commit,
  read,
  subscribe,
}: PainterTargetOptions<T>): Promise<T> =>
  new Promise((resolve, reject) => {
    let phase: "committing" | "ready" | "settled" = "committing";
    const check = () => {
      if (phase !== "ready") return;
      const result = Result.try({
        try: () => ({ target: read() }),
        catch: (error: unknown) => error,
      });
      if (result.isErr()) {
        phase = "settled";
        unsubscribe();
        reject(result.error);
        return;
      }
      if (result.value.target === null) return;
      phase = "settled";
      unsubscribe();
      resolve(result.value.target);
    };
    const unsubscribe = subscribe(check);
    const committed = Result.try({ try: commit, catch: (error: unknown) => error });
    if (committed.isErr()) {
      phase = "settled";
      unsubscribe();
      reject(committed.error);
      return;
    }
    phase = "ready";
    check();
  });
