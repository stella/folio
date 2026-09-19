/**
 * The closed vocabulary a difference signature is allowed to quote.
 *
 * A signature that inlines a value read from a document is wrong twice over:
 * third-party content lands in a committed baseline, which `corpus/README.md`
 * forbids outright, and the row moves whenever that document's text changes, so
 * one defect fans out into a row per distinct value pair. The formatter
 * therefore quotes a string only when the string is a token the format itself
 * defines, and reports every other string by shape.
 *
 * The vocabulary is derived, never hand-listed, from the two places that own
 * the closed sets the invariants compare:
 *
 * - every `enumValues` member of every enumerated simple type in the committed
 *   schema graph (`specifications/generated/docx-transitional-schema.gen.json`),
 *   which is where `strict`/`transitional`, `restart`/`continue`, `rtl`/`ltr`
 *   and every other `ST_*` token is declared;
 * - every string value of every `as const` record `@stll/docx-core/model`
 *   exports, which is where folio's own sets live: `DOCX_CONFORMANCE_CLASSES`
 *   adds `unknown` to the schema's two conformance classes, `BIDI_CONTROLS`,
 *   `DRAWING_RAW_XML_MODES`, `PARSE_WARNING_CODES` and `REVIEW_CARRIERS` have
 *   no schema counterpart at all.
 *
 * A token added to either source joins the vocabulary on the next run. A value
 * in neither is document content by elimination, and the formatter reports its
 * shape instead.
 */

import { readFileSync } from "node:fs";
import path from "node:path";

// By path, not by specifier: the gate's scripts are not a workspace package,
// and `@stll/docx-core` is reachable only from the packages that depend on it.
import * as docxModel from "../../../packages/docx-core/src/model/document";

const GRAPH_PATH = path.join(
  path.resolve(import.meta.dir, "../../.."),
  "specifications/generated/docx-transitional-schema.gen.json",
);

/**
 * Only the shape the vocabulary reads, so the graph's full type stays out of a
 * hot path that never looks at anything else.
 */
type EnumeratedSymbols = {
  symbols: ReadonlyArray<{ kind: string; enumValues?: readonly string[] }>;
};

/**
 * The longest token either source declares is 28 characters, so anything past
 * this is not one: the bound stops a value that merely collides with a token
 * name from carrying a document's text into a row.
 */
const MAX_TOKEN_LENGTH = 32;

const isToken = (value: string): boolean =>
  value.length > 0 && value.length <= MAX_TOKEN_LENGTH && !/\s/u.test(value);

const schemaEnumValues = (): string[] => {
  // Read synchronously and once: the formatter is called from inside a
  // synchronous model walk, and the parse costs single-digit milliseconds per
  // worker process against a run measured in minutes.
  const graph = JSON.parse(readFileSync(GRAPH_PATH, "utf8")) as EnumeratedSymbols;
  const values: string[] = [];
  for (const symbol of graph.symbols) {
    if (symbol.kind !== "simpleType" || symbol.enumValues === undefined) {
      continue;
    }
    values.push(...symbol.enumValues);
  }
  return values;
};

const isStringRecord = (value: unknown): value is Record<string, string> =>
  typeof value === "object" &&
  value !== null &&
  !Array.isArray(value) &&
  Object.values(value).length > 0 &&
  Object.values(value).every((entry) => typeof entry === "string");

const modelConstantValues = (): string[] => {
  const values: string[] = [];
  for (const exported of Object.values(docxModel)) {
    if (isStringRecord(exported)) {
      values.push(...Object.values(exported));
    }
  }
  return values;
};

let vocabulary: ReadonlySet<string> | undefined;

/** Every token a signature may name, derived once per process. */
export const valueVocabulary = (): ReadonlySet<string> => {
  vocabulary ??= new Set([...schemaEnumValues(), ...modelConstantValues()].filter(isToken));
  return vocabulary;
};

export const isVocabularyToken = (value: string): boolean => valueVocabulary().has(value);
