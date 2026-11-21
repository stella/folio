import type { Node as ProseMirrorNode } from "prosemirror-model";

/** Which ends of an edit stand just inside a content control's edge. */
export type ControlEdges = { start?: true; end?: true };

/**
 * Just inside a content control's edge, the editor position, not the source
 * gap, says the text belongs to the control: a replacement of exactly its
 * text goes in it, and the control stays.
 */
export const controlEdgesAt = (
  doc: ProseMirrorNode,
  from: number,
  to: number,
): ControlEdges => {
  const innerEdge = (position: number, side: "start" | "end") => {
    const resolved = doc.resolve(position);
    return (
      resolved.parent.type.name === "sdt" &&
      resolved.parentOffset ===
        (side === "start" ? 0 : resolved.parent.content.size)
    );
  };
  return {
    ...(innerEdge(from, "start") ? { start: true } : {}),
    ...(innerEdge(to, "end") ? { end: true } : {}),
  };
};
