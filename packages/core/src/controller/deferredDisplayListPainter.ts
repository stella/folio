import { Result } from "better-result";

import type { DisplayListPagePainterFactory } from "./layoutPipelineExecution";

let painter: DisplayListPagePainterFactory | undefined;
let pending: Promise<Result<DisplayListPagePainterFactory, Error>> | undefined;

const loadPainter = () => {
  pending ??= Result.tryPromise({
    try: async () => {
      const { createDisplayListPagePainter } =
        await import("../display-list/editor/displayListPagePainter");
      painter = createDisplayListPagePainter;
      return painter;
    },
    catch: (cause) =>
      new Error("Could not load the display-list painter.", { cause }),
  }).then((result) => {
    if (result.isErr()) pending = undefined;
    return result;
  });
  return pending;
};

type DisplayListPainterLoaderOptions = {
  onReady: () => void;
  onError: (error: Error) => void;
};

/** Notify an editor once per in-flight load, regardless of how many layouts requested it. */
export const createDisplayListPainterLoader = ({
  onReady,
  onError,
}: DisplayListPainterLoaderOptions) => {
  let waiting = false;
  return (): DisplayListPagePainterFactory | undefined => {
    if (painter) return painter;
    if (!waiting) {
      waiting = true;
      void loadPainter().then((result) => {
        waiting = false;
        if (result.isOk()) onReady();
        else onError(result.error);
        return undefined;
      });
    }
    return undefined;
  };
};
