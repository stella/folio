import { Result } from "better-result";

type PainterTargetOptions<T> = {
  read: () => T | null;
  subscribe: (listener: () => void) => () => void;
};

/** Subscribe before reading, then resolve a fresh target from the successful paint. */
export const resolvePainterTarget = <T>({ read, subscribe }: PainterTargetOptions<T>): Promise<T> =>
  new Promise((resolve, reject) => {
    const check = () => {
      const result = Result.try({
        try: () => ({ target: read() }),
        catch: (error: unknown) => error,
      });
      if (result.isErr()) {
        unsubscribe();
        reject(result.error);
        return;
      }
      if (result.value.target === null) return;
      unsubscribe();
      resolve(result.value.target);
    };
    const unsubscribe = subscribe(check);
    check();
  });
