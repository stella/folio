import type { Node as PMNode } from "prosemirror-model";

import type { EditorIntent } from "@stll/docx-core/ops";

import type { TextChange } from "../ai-edits/minimal-replacement";
import type { CanonicalProjection } from "./canonicalSession";

type Sources = NonNullable<
  Extract<EditorIntent, { type: "replaceText" }>["formattingSources"]
>;
type PlacedChange = TextChange & { from: number; to: number };

type ReplacementRunSourcesOptions = {
  paragraph: PMNode;
  paragraphStart: number;
  offsets: readonly number[];
  start: number;
  end: number;
  direct: readonly PlacedChange[];
  tracked: readonly PlacedChange[];
  projection: Pick<CanonicalProjection, "addressAt">;
};

/**
 * Both plans write the same replacement. For every tracked insertion, read
 * its run sources from the corresponding output span of the direct plan.
 * Text the direct plan keeps retains its own run, and new text inherits at
 * that direct edit's range or caret.
 */
export const replacementRunSources = ({
  paragraph,
  paragraphStart,
  offsets,
  start,
  end,
  direct,
  tracked,
  projection,
}: ReplacementRunSourcesOptions): (Sources | undefined)[] | undefined => {
  if (
    direct.length === tracked.length &&
    direct.every((change, index) => {
      const other = tracked[index];
      return (
        other?.start === change.start &&
        other.end === change.end &&
        other.text === change.text
      );
    })
  )
    return undefined;

  const parts: Sources[number][] = [];
  let target = 0;
  let consumed = start;
  const keepUntil = (limit: number) => {
    let run: PMNode | null = null;
    for (let offset = consumed; offset < limit; offset++) {
      const position = offsets[offset];
      if (position === undefined) return false;
      const next = paragraph.nodeAt(position - paragraphStart);
      target++;
      const previous = parts.at(-1);
      if (run !== null && next === run && previous !== undefined) {
        previous.end = target;
        continue;
      }
      const at = projection.addressAt(position, "afterMarkers");
      if (at.isErr()) return false;
      parts.push({ type: "character", at: at.value, end: target });
      run = next;
    }
    return true;
  };
  for (const change of direct) {
    if (!keepUntil(change.start)) return undefined;
    consumed = change.end;
    if (change.text === "") continue;
    const caret = change.from === change.to;
    const at = projection.addressAt(
      change.from,
      caret ? undefined : "afterMarkers",
    );
    if (at.isErr()) return undefined;
    target += change.text.length;
    parts.push({
      type: caret ? "caret" : "character",
      at: at.value,
      end: target,
    });
  }
  if (!keepUntil(end)) return undefined;

  let shift = -start;
  let partIndex = 0;
  return tracked.map((change) => {
    const from = change.start + shift;
    const to = from + change.text.length;
    shift += change.text.length - (change.end - change.start);
    if (from === to) return undefined;
    let part = parts[partIndex];
    while (part !== undefined && part.end <= from) {
      partIndex++;
      part = parts[partIndex];
    }
    const runs: Sources[number][] = [];
    while (part !== undefined) {
      runs.push({ ...part, end: Math.min(to, part.end) - from });
      if (part.end >= to) return runs;
      partIndex++;
      part = parts[partIndex];
    }
    return undefined;
  });
};
