import type { Node as PMNode } from "prosemirror-model";
import type { NoteMarkerAttrs } from "../prosemirror/schema/nodes";
import {
  attrsRecord,
  attrsResult,
  expectCachedNodeAttrs,
  expectNodeType,
  type ProseMirrorAttrIssue,
  type ReadProseMirrorAttrsResult,
} from "./prosemirrorAttrBoundary";

const noteMarkerAttrsCache = new WeakMap<PMNode, NoteMarkerAttrs>();

export const readNoteMarkerAttrs = (node: PMNode): ReadProseMirrorAttrsResult<NoteMarkerAttrs> => {
  const attrs = attrsRecord(node.attrs);
  const issues: ProseMirrorAttrIssue[] = [];
  expectNodeType(node, "noteMarker", issues);
  if (attrs["kind"] !== "footnote" && attrs["kind"] !== "endnote") {
    issues.push({ path: "noteMarker.attrs.kind", message: "Expected footnote or endnote." });
  }
  return attrsResult(attrs, issues);
};

export const expectNoteMarkerAttrs = (node: PMNode): NoteMarkerAttrs =>
  expectCachedNodeAttrs(node, noteMarkerAttrsCache, readNoteMarkerAttrs, "noteMarker attrs");
