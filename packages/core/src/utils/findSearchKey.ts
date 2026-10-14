import { documentStories, storyBody } from "@stll/docx-core/ops";
import type { Document } from "../types/document";
import type { FindOptions, FindSearchContext } from "./findReplace";

/** Every matching option must make an explicit decision in the search key. */
export const createFindQueryKey = (query: string, options: FindOptions): string => {
  const matchingOptions = {
    matchCase: options.matchCase,
    matchWholeWord: options.matchWholeWord,
    useRegex: options.useRegex ?? false,
    displayMode: options.displayMode ?? "all-markup",
    includeComments: options.includeComments ?? false,
  } satisfies Required<FindOptions>;
  return JSON.stringify([query, matchingOptions]);
};

type FindSearchKeyOptions = {
  query: string;
  options: FindOptions;
  document: Document;
  /** Immutable editor document identity, changed by every body transaction. */
  bodyVersion?: object;
  context?: FindSearchContext;
};

/**
 * Getters can return fresh model clones. Keep the immutable body version and
 * snapshot package stories separately: secondary edits do not change that body.
 * Story enumeration also binds this snapshot to the package's story resolver.
 */
export const createFindSearchKey = ({
  query,
  options,
  document,
  bodyVersion,
  context,
}: FindSearchKeyOptions) => ({
  query: createFindQueryKey(query, options),
  bodyVersion: bodyVersion ?? JSON.stringify(document.package.document.content),
  documentVersion: JSON.stringify([
    documentStories(document)
      .filter((story) => story !== "main")
      .map((story) => [story, storyBody(document, story)]),
    document.package.document.comments ?? [],
    document.package.document.sections?.map(({ properties }) => properties) ?? [],
  ]),
  storyScope: JSON.stringify(
    context?.renderedStories?.map((story) => JSON.stringify(story)).toSorted() ?? null,
  ),
});

export const sameFindSearchKey = (
  previous: ReturnType<typeof createFindSearchKey> | null,
  current: ReturnType<typeof createFindSearchKey>,
): boolean =>
  previous !== null &&
  previous.query === current.query &&
  sameFindDocumentVersion(previous, current) &&
  previous.storyScope === current.storyScope;

const sameFindDocumentVersion = (
  previous: ReturnType<typeof createFindSearchKey>,
  current: ReturnType<typeof createFindSearchKey>,
): boolean =>
  sameFindBodyVersion(previous, current) && previous.documentVersion === current.documentVersion;

/** Secondary edits cannot acknowledge a pending body replacement. */
export const sameFindBodyVersion = (
  previous: ReturnType<typeof createFindSearchKey>,
  current: ReturnType<typeof createFindSearchKey>,
): boolean => previous.bodyVersion === current.bodyVersion;
