/**
 * `folio placeholders`: find the placeholders a draft still carries, in the
 * body and every other story, and fill them from a map of values. A fill
 * replaces each placeholder's exact range, as a tracked change unless the
 * caller asks for a direct edit, so the run formatting around it stays.
 *
 * A placeholder is a match of one of the patterns. A pattern's `name` group
 * (or its first group) names the field that fills it (`{{party}}` is filled
 * by `party`); any placeholder is also filled by a value keyed by its exact
 * text (`[•]`).
 */

import { Result } from "better-result";

import {
  createFolioAITextRangeHandle,
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioAIEditApplyMode,
  type FolioDocumentStoryHandle,
  type FolioDocxReviewer,
} from "@stll/folio-core/server";
import { sliceText } from "@stll/folio-core/utils/sliceText";

import { cliError, FOLIO_CLI_ERROR_CODES, type FolioCliError } from "./errors";

/** The patterns a draft's placeholders usually take. */
export const DEFAULT_PLACEHOLDER_PATTERNS: readonly string[] = [
  String.raw`\[•\]`,
  String.raw`\[…\]`,
  String.raw`\[\.\.\.\]`,
  String.raw`\{\{\s*(?<name>[^{}]+?)\s*\}\}`,
  String.raw`\[\[\s*(?<name>[^[\]]+?)\s*\]\]`,
  String.raw`\bTODO\b`,
];

/** Characters of context shown on each side of a placeholder. */
const CONTEXT_CHARACTERS = 40;

/** One placeholder: where it is, what it says, and what fills it. */
export type Placeholder = {
  story: FolioDocumentStoryHandle;
  blockId: string;
  text: string;
  /** The field it names, when its pattern names one. */
  name?: string;
  startOffset: number;
  endOffset: number;
  context: string;
  /** The value that fills it, when the fill map has one. */
  fill?: string;
};

export type PlaceholderReport = {
  /** Placeholders found (before a fill) or left (after one). */
  placeholders: Placeholder[];
  /** Placeholders a fill replaced; 0 when nothing was written. */
  filled: number;
  /** Placeholders no value fills (after a fill: those left). */
  unresolved: number;
};

type PlaceholderOptions = {
  patterns: readonly RegExp[];
  fields: ReadonlyMap<string, string>;
};

const invalidInput = (message: string, hint?: string): FolioCliError =>
  cliError({ code: FOLIO_CLI_ERROR_CODES.invalidInput, message, hint });

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const compilePatterns = (
  value: unknown,
): Result<readonly RegExp[], FolioCliError> => {
  if (value !== undefined && !Array.isArray(value)) {
    return Result.err(invalidInput("patterns must be a list of expressions."));
  }
  const sources: readonly unknown[] = value ?? DEFAULT_PLACEHOLDER_PATTERNS;
  if (sources.length === 0) {
    return Result.err(invalidInput("patterns must name at least one."));
  }
  const patterns: RegExp[] = [];
  for (const source of sources) {
    if (typeof source !== "string" || source === "") {
      return Result.err(
        invalidInput("Every pattern must be a non-empty string."),
      );
    }
    const compiled = Result.try(() => new RegExp(source, "gu"));
    if (compiled.isErr()) {
      return Result.err(
        invalidInput(
          `Pattern ${JSON.stringify(source)} is not a valid expression.`,
        ),
      );
    }
    if (new RegExp(source, "u").test("")) {
      return Result.err(
        invalidInput(`Pattern ${JSON.stringify(source)} matches empty text.`),
      );
    }
    patterns.push(compiled.value);
  }
  return Result.ok(patterns);
};

const readFields = (
  value: unknown,
): Result<ReadonlyMap<string, string>, FolioCliError> => {
  if (value === undefined) return Result.ok(new Map());
  if (!isRecord(value)) {
    return Result.err(
      invalidInput(
        "fill must be an object of field values.",
        'For example {"party": "Acme Ltd", "[•]": "1 January 2027"}.',
      ),
    );
  }
  const fields = new Map<string, string>();
  for (const [key, entry] of Object.entries(value)) {
    if (typeof entry !== "string") {
      return Result.err(
        invalidInput(`The value for ${JSON.stringify(key)} must be a string.`),
      );
    }
    fields.set(key, entry);
  }
  return Result.ok(fields);
};

const optionsFrom = (
  args: Readonly<Record<string, unknown>>,
): Result<PlaceholderOptions, FolioCliError> => {
  const patterns = compilePatterns(args["patterns"]);
  if (patterns.isErr()) return Result.err(patterns.error);
  const fields = readFields(args["fill"]);
  if (fields.isErr()) return Result.err(fields.error);
  for (const [key, value] of fields.value) {
    // A value that is itself a placeholder would be found again after the fill.
    if (
      patterns.value.some((pattern) =>
        new RegExp(pattern.source, "u").test(value),
      )
    ) {
      return Result.err(
        invalidInput(
          `The value for ${JSON.stringify(key)} contains a placeholder.`,
        ),
      );
    }
  }
  return Result.ok({ patterns: patterns.value, fields: fields.value });
};

const contextOf = (text: string, start: number, end: number): string => {
  const context = sliceText({
    text,
    units: "codeUnits",
    start: Math.max(0, start - CONTEXT_CHARACTERS),
    end: end + CONTEXT_CHARACTERS,
  });
  return `${start > CONTEXT_CHARACTERS ? "…" : ""}${context}${end + CONTEXT_CHARACTERS < text.length ? "…" : ""}`;
};

type BlockLocation = Pick<Placeholder, "story" | "blockId">;

/** The placeholders in one block's text: first match wins where two overlap. */
const matchesIn = (
  text: string,
  options: PlaceholderOptions,
  { story, blockId }: BlockLocation,
): Placeholder[] => {
  const found: Placeholder[] = [];
  for (const pattern of options.patterns) {
    for (const match of text.matchAll(pattern)) {
      const startOffset = match.index;
      const endOffset = startOffset + match[0].length;
      const name = (match.groups?.["name"] ?? match[1])?.trim();
      const fill =
        (name === undefined ? undefined : options.fields.get(name)) ??
        options.fields.get(match[0]);
      found.push({
        story,
        blockId,
        text: match[0],
        ...(name !== undefined && name !== "" && { name }),
        startOffset,
        endOffset,
        context: contextOf(text, startOffset, endOffset),
        ...(fill !== undefined && { fill }),
      });
    }
  }
  const kept: typeof found = [];
  for (const match of found.toSorted(
    (left, right) => left.startOffset - right.startOffset,
  )) {
    const previous = kept.at(-1);
    if (previous === undefined || match.startOffset >= previous.endOffset)
      kept.push(match);
  }
  return kept;
};

type StoryBlocks = {
  story: FolioDocumentStoryHandle;
  blocks: readonly { id: string; text: string }[];
};

/** Every story's text blocks, the body first. */
const storyBlocks = (reviewer: FolioDocxReviewer): StoryBlocks[] =>
  reviewer.listStories().flatMap(({ handle }) => {
    const snapshot = reviewer.snapshotStory(handle);
    if (snapshot === null) return [];
    return [
      {
        story: handle,
        blocks: snapshot.blocks.flatMap((block) =>
          block.kind === "diagnostic"
            ? []
            : [{ id: block.id, text: block.text }],
        ),
      },
    ];
  });

const findPlaceholders = (
  reviewer: FolioDocxReviewer,
  options: PlaceholderOptions,
): Placeholder[] =>
  storyBlocks(reviewer).flatMap(({ story, blocks }) =>
    blocks.flatMap(({ id, text }) =>
      matchesIn(text, options, { story, blockId: id }),
    ),
  );

/** List the placeholders and the values that would fill them; writes nothing. */
export const inspectPlaceholders = (
  reviewer: FolioDocxReviewer,
  args: Readonly<Record<string, unknown>>,
): Result<PlaceholderReport, FolioCliError> => {
  const options = optionsFrom(args);
  if (options.isErr()) return Result.err(options.error);
  const placeholders = findPlaceholders(reviewer, options.value);
  return Result.ok({
    placeholders,
    filled: 0,
    unresolved: placeholders.filter(({ fill }) => fill === undefined).length,
  });
};

export type FillContext = {
  mode: FolioAIEditApplyMode;
  date: string;
  idSeed: number;
};

const sameStory = (
  left: FolioDocumentStoryHandle,
  right: FolioDocumentStoryHandle,
): boolean => JSON.stringify(left) === JSON.stringify(right);

const refused = (reasons: unknown): FolioCliError =>
  cliError({
    code: FOLIO_CLI_ERROR_CODES.operationRejected,
    message: "A placeholder could not be filled; nothing was written.",
    details: { skipped: reasons },
  });

/**
 * Fill every placeholder the map has a value for. Each pass replaces the
 * first fillable placeholder of every block, story by story, and reads the
 * stories again, so ranges always come from the text they edit.
 */
export const fillPlaceholders = (
  reviewer: FolioDocxReviewer,
  args: Readonly<Record<string, unknown>>,
  context: FillContext,
): Result<PlaceholderReport, FolioCliError> => {
  const options = optionsFrom(args);
  if (options.isErr()) return Result.err(options.error);
  if (options.value.fields.size === 0) {
    return Result.err(invalidInput("fill names no values to write."));
  }
  let idSeed = context.idSeed;
  let filled = 0;
  const limit = findPlaceholders(reviewer, options.value).length;
  for (let pass = 0; pass <= limit; pass += 1) {
    const fillable = findPlaceholders(reviewer, options.value).filter(
      (placeholder) => placeholder.fill !== undefined,
    );
    if (fillable.length === 0) break;
    const stories: FolioDocumentStoryHandle[] = [];
    for (const { story } of fillable) {
      if (!stories.some((known) => sameStory(known, story)))
        stories.push(story);
    }
    for (const story of stories) {
      const snapshot = reviewer.snapshotStory(story);
      if (snapshot === null) continue;
      const seen = new Set<string>();
      const operations = [];
      for (const placeholder of fillable) {
        if (
          !sameStory(placeholder.story, story) ||
          seen.has(placeholder.blockId)
        )
          continue;
        seen.add(placeholder.blockId);
        const block = snapshot.blocks.find(
          ({ id }) => id === placeholder.blockId,
        );
        if (block === undefined || block.kind === "diagnostic") continue;
        const range = createFolioAITextRangeHandle({
          blockId: block.id,
          text: block.text,
          startOffset: placeholder.startOffset,
          endOffset: placeholder.endOffset,
        });
        if (range === null) continue;
        operations.push({
          id: `placeholder-${String(pass)}-${String(operations.length)}`,
          type: "replaceRange" as const,
          range,
          replace: placeholder.fill ?? "",
        });
      }
      if (operations.length === 0) continue;
      const result = reviewer.applyDocumentOperationsToStory({
        story,
        batch: {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode: context.mode,
          operations,
        },
        snapshot,
        revisionStamp: { date: context.date, idSeed },
        undefinedReferences: "refuse",
      });
      if (result.skipped.length > 0) return Result.err(refused(result.skipped));
      filled += result.applied.length;
      idSeed = Math.max(idSeed, result.nextRevisionId ?? idSeed);
    }
  }
  const placeholders = findPlaceholders(reviewer, options.value);
  return Result.ok({
    placeholders,
    filled,
    unresolved: placeholders.length,
  });
};
