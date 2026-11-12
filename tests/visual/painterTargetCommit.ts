import { Result } from "better-result";

/** What painted-target readiness reads of the page's font loading. */
export type PainterFontSet = Pick<FontFaceSet, "status" | "ready">;

type PainterTargetOptions<T> = {
  commit: () => void;
  read: () => T | null;
  subscribe: (listener: () => void) => () => void;
  /** The page's font set; `document.fonts` unless a test supplies one. */
  fonts?: PainterFontSet | null;
};

const pageFonts = (): PainterFontSet | null =>
  typeof document === "undefined" ? null : document.fonts;

const loading = (fonts: PainterFontSet | null): fonts is PainterFontSet =>
  fonts !== null && fonts.status === "loading";

/** Commit current document paint before reading; numeric hit tests can match stale pages. */
const committedRead = <T>({
  commit,
  read,
  subscribe,
}: Omit<PainterTargetOptions<T>, "fonts">): Promise<T> =>
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
    const committed = Result.try({
      try: commit,
      catch: (error: unknown) => error,
    });
    if (committed.isErr()) {
      phase = "settled";
      unsubscribe();
      reject(committed.error);
      return;
    }
    phase = "ready";
    check();
  });

/**
 * A target read from paint that no font load can move.
 *
 * The editor measures text in the faces loaded when it lays out, and lays out
 * again when a face that changes a width loads later. Painting text can start
 * loading the `unicode-range` subsets it needs, so coordinates read while a
 * face loads name glyphs that the next layout moves, and a drag lands one
 * position off. So the paint is committed with nothing loading, the target is
 * read from it, and the read stands only when the layout it forces starts no
 * load; otherwise the fonts settle and the paint is committed and read again.
 */
export const resolvePainterTarget = <T>({
  fonts = pageFonts(),
  ...options
}: PainterTargetOptions<T>): Promise<T> => {
  // Settled fonts commit at once, so the subscription exists when this returns.
  const attempt = (): Promise<T> =>
    loading(fonts)
      ? fonts.ready.then(attempt)
      : committedRead(options).then((target) => {
          // Style and layout start the loads the painted text needs.
          if (typeof document !== "undefined")
            document.documentElement.getBoundingClientRect();
          return loading(fonts) ? attempt() : target;
        });
  return attempt();
};
