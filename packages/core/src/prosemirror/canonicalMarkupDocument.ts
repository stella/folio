import type { Node as PMNode } from "prosemirror-model";
import type { EditorState } from "prosemirror-state";
import type { StepMap } from "prosemirror-transform";

import {
  findStoryBody,
  OP_STORIES,
  resolvedMarkupDocumentView,
  type RevisionDecision,
  sameStory,
} from "@stll/docx-core/ops";

import type { Document, StyleDefinitions } from "../types/document";
import { writableCopy } from "../types/documentModel";
import { proseDocToBlocks } from "./conversion/fromProseDoc";
import {
  markupViewSources,
  markupViewStateSources,
} from "./conversion/markupViewSource";
import { storyConversionOptions } from "./conversion/storyConversionOptions";
import {
  footnoteToProseDoc,
  headerFooterToProseDoc,
  toProseDoc,
} from "./conversion/toProseDoc";
import {
  projectDetachedText,
  projectIndependentMarkup,
} from "./independentMarkupProjection";
import { getStatedDocumentNumbering } from "./plugins/documentNumbering";
import { getDocumentStyleResolver } from "./plugins/documentStyleState";
import { styleResolverSource } from "./styles/styleResolver";

type CanonicalMarkupDocumentOptions = {
  state: EditorState;
  decision: RevisionDecision;
};

const detachedModel = (state: EditorState): Document => {
  const numbering = getStatedDocumentNumbering(state);
  const resolver = getDocumentStyleResolver(state);
  const definitions =
    resolver === null ? undefined : styleResolverSource(resolver);
  const styles =
    definitions === undefined
      ? undefined
      : writableCopy<StyleDefinitions>(definitions);
  return {
    package: {
      document: { content: proseDocToBlocks(state.doc, undefined, styles) },
      ...(numbering === null ? {} : { numbering }),
      ...(styles === undefined ? {} : { styles }),
    },
  };
};

type CanonicalMarkupDocument = {
  doc: PMNode;
  partial: boolean;
  positionMap?: StepMap;
  changeBarRanges?: readonly { from: number; to: number }[];
};

/** Resolve a copy of the authored model and project the selected review view. */
export const canonicalMarkupDocument = ({
  state,
  decision,
}: CanonicalMarkupDocumentOptions): CanonicalMarkupDocument => {
  const converted =
    markupViewStateSources.get(state) ?? markupViewSources.get(state.doc);
  const authoritative = converted?.document;
  if (authoritative === undefined) {
    const detached = projectDetachedText(state.doc, decision);
    if (detached !== undefined) return { ...detached, partial: false };
  }
  const source = authoritative ?? detachedModel(state);
  const resolved = resolvedMarkupDocumentView(source, decision);
  const story = converted?.story ?? OP_STORIES.MAIN;
  const partial = resolved.unresolved.some((unresolved) =>
    sameStory(unresolved, story),
  );
  if (resolved.document === source) return { doc: state.doc, partial };
  if (story === OP_STORIES.MAIN) {
    const projected =
      resolved.independentInline === undefined
        ? undefined
        : projectIndependentMarkup({
            doc: state.doc,
            changes: resolved.independentInline,
          });
    if (projected !== undefined) return { ...projected, partial };
    const projectedDoc = toProseDoc(resolved.document, converted?.options);
    return {
      doc:
        authoritative === undefined
          ? state.doc.type.create(state.doc.attrs, projectedDoc.content)
          : projectedDoc,
      partial,
    };
  }
  const body = findStoryBody(resolved.document, story);
  if (body === undefined) return { doc: state.doc, partial: true };
  const {
    styles: storyStyles,
    theme,
    numbering: storyNumbering,
  } = resolved.document.package;
  const options =
    converted?.options ??
    storyConversionOptions({
      styles: storyStyles,
      theme,
      numbering: storyNumbering,
    });
  return {
    doc:
      story.kind === "header" || story.kind === "footer"
        ? headerFooterToProseDoc(body.content, options)
        : footnoteToProseDoc(body.content, options),
    partial,
  };
};
