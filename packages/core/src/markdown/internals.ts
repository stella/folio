/**
 * Internal helpers shared by the markdown converter: context construction and
 * warning dedupe. Kept free of renderer imports so the renderers can import
 * `pushWarning` without forming an import cycle (trailer emission, which needs
 * `renderBlock`, lives in `trailers.ts`). Ported from eigenpal/docx-editor
 * PR #595, trimmed to the sync continuous path.
 */

import { createBuiltInStyleIndex } from "../docx/builtInStyles";
import { createListLabelCounter } from "../prosemirror/listLabels";
import type { StyleDefinitions } from "../types/document";
import {
  createNoteReferenceNumbering,
  noteReferenceMarker,
  noteReferenceToken,
  type NoteReferenceKind,
} from "../utils/noteReferenceLabels";
import type { MarkdownOptions, RenderContext } from "./types";

/**
 * Build a fresh `RenderContext` from caller options, applying defaults. The
 * `footnotes` default is `"keep"` (folio addition over upstream #595).
 *
 * The document's styles are indexed once here: every paragraph asks the index
 * whether it is a heading or a quote, so rebuilding it per paragraph would
 * make the render quadratic.
 */
export function newContext(
  opts: MarkdownOptions = {},
  styles?: StyleDefinitions | undefined,
): RenderContext {
  return {
    builtInStyles: createBuiltInStyleIndex(styles?.styles ?? [], styles?.docDefaults),
    opts: {
      annotations: opts.annotations ?? "html",
      trackedChanges: opts.trackedChanges ?? "annotate",
      comments: opts.comments ?? "inline",
      hyperlinks: opts.hyperlinks ?? "inline",
      footnotes: opts.footnotes ?? "keep",
      imagePath: opts.imagePath,
    },
    images: new Map(),
    imagesByPath: new Map(),
    warnings: [],
    footnoteRefs: [],
    noteNumbering: createNoteReferenceNumbering(),
    commentRefs: [],
    hyperlinkRefs: [],
    imageCounter: 0,
    nextListLabel: createListLabelCounter(),
  };
}

/**
 * Push a warning into the context, deduplicating against existing entries so
 * recurring messages appear at most once.
 */
export function pushWarning(ctx: RenderContext, message: string): void {
  if (!ctx.warnings.includes(message)) {
    ctx.warnings.push(message);
  }
}

/**
 * Number a footnote or endnote reference in reading order and record its
 * note for the definitions trailer on the note's first reference. Returns the
 * bare token (`1`, `e1`) and the inline marker (`[^1]`, `[^e1]`).
 */
export function numberNoteReference(
  ctx: RenderContext,
  kind: NoteReferenceKind,
  refId: number,
): { token: string; marker: string } {
  const { displayNumber, first } = ctx.noteNumbering.next(kind, refId);
  const marker = noteReferenceMarker(kind, displayNumber);
  if (first) {
    ctx.footnoteRefs.push({ refId, marker, kind });
  }
  return { token: noteReferenceToken(kind, displayNumber), marker };
}
