import type { Node as PMNode } from "prosemirror-model";
import type { EditorState } from "prosemirror-state";

import type { OpStory } from "@stll/docx-core/ops";

import type { Document } from "../../types/document";
import type { ToProseDocOptions } from "./toProseDoc";

type MarkupViewSource = {
  document: Document;
  story?: OpStory;
  options?: ToProseDocOptions;
};

/** The model used to convert a story, keyed by the exact PM snapshot. */
export const markupViewSources = new WeakMap<PMNode, MarkupViewSource>();

/** Two states can share a PM node while publishing different model snapshots. */
export const markupViewStateSources = new WeakMap<
  EditorState,
  MarkupViewSource
>();

/** Carry the published model through a transaction that preserves the story. */
export const retainMarkupViewSource = (
  before: EditorState,
  after: EditorState,
): void => {
  const source =
    markupViewStateSources.get(before) ?? markupViewSources.get(before.doc);
  if (source !== undefined) markupViewStateSources.set(after, source);
};
