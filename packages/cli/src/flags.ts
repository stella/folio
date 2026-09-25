/**
 * Command flags generated from a tool's JSON Schema. Each top-level property
 * becomes `--kebab-case-name`: booleans are switches, numbers and strings take
 * a value, arrays of strings repeat (`--id 3 --id 4`, or `--id 3,4`), and
 * other objects or arrays take a JSON value. `--input` supplies the
 * whole argument object as JSON (inline, `@file`, or `-` for stdin), and
 * flags given alongside it override its fields.
 */

import { panic, Result } from "better-result";
import { readFile } from "node:fs/promises";

import { cliError, FOLIO_CLI_ERROR_CODES, type FolioCliError } from "./errors";
import type { JsonObjectSchema } from "./registry";

/** Largest `--input` document accepted. */
export const MAX_INPUT_BYTES = 16 * 1024 * 1024;

export type FlagValueType = "boolean" | "integer" | "number" | "string" | "strings" | "json";

export type GeneratedFlag = {
  /** The schema property the flag sets. */
  property: string;
  /** `--flag` spelling, without the dashes. */
  flag: string;
  valueType: FlagValueType;
  description: string;
};

export const kebabCase = (name: string): string =>
  name.replace(/[A-Z]/gu, (letter) => `-${letter.toLowerCase()}`);

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const valueTypeOf = (schema: unknown): FlagValueType => {
  if (!isRecord(schema)) return "json";
  switch (schema["type"]) {
    case "boolean":
      return "boolean";
    case "integer":
      return "integer";
    case "number":
      return "number";
    case "string":
      return "string";
    case "array": {
      const items = schema["items"];
      return isRecord(items) && items["type"] === "string" ? "strings" : "json";
    }
    default:
      return "json";
  }
};

const descriptionOf = (schema: unknown): string => {
  if (!isRecord(schema)) return "";
  const description = typeof schema["description"] === "string" ? schema["description"] : "";
  const values = Array.isArray(schema["enum"]) ? ` One of: ${schema["enum"].join(", ")}.` : "";
  return `${description}${values}`.trim();
};

/**
 * One flag per top-level schema property, in declaration order. `flagNames`
 * renames a property's flag (`{ ids: "id" }` for `--id`).
 */
export const flagsForSchema = (
  schema: JsonObjectSchema,
  flagNames: Readonly<Record<string, string>> = {},
): GeneratedFlag[] =>
  Object.entries(schema.properties).map(([property, propertySchema]) => ({
    property,
    flag: flagNames[property] ?? kebabCase(property),
    valueType: valueTypeOf(propertySchema),
    description: descriptionOf(propertySchema),
  }));

const usage = (message: string, hint?: string): FolioCliError =>
  cliError({ code: FOLIO_CLI_ERROR_CODES.usage, message, hint });

type CoerceFlagOptions = { flag: GeneratedFlag; raw: string | boolean | readonly string[] };

/** Convert one parsed flag value to the type its schema property declares. */
export const coerceFlag = ({ flag, raw }: CoerceFlagOptions): Result<unknown, FolioCliError> => {
  if (typeof raw === "boolean") {
    return flag.valueType === "boolean"
      ? Result.ok(raw)
      : Result.err(usage(`--${flag.flag} takes a value.`));
  }
  if (typeof raw !== "string") {
    return flag.valueType === "strings"
      ? Result.ok(raw.flatMap((value) => value.split(",")).filter((value) => value !== ""))
      : Result.err(usage(`--${flag.flag} is given more than once.`));
  }
  switch (flag.valueType) {
    case "boolean":
      return Result.err(usage(`--${flag.flag} is a switch and takes no value.`));
    case "string":
      return Result.ok(raw);
    case "strings":
      return Result.ok(raw.split(",").filter((value) => value !== ""));
    case "integer":
    case "number": {
      const value = Number(raw);
      if (raw.trim() === "" || !Number.isFinite(value)) {
        return Result.err(usage(`--${flag.flag} expects a number, got ${JSON.stringify(raw)}.`));
      }
      if (flag.valueType === "integer" && !Number.isInteger(value)) {
        return Result.err(usage(`--${flag.flag} expects an integer, got ${JSON.stringify(raw)}.`));
      }
      return Result.ok(value);
    }
    case "json": {
      const parsed = Result.try((): unknown => JSON.parse(raw));
      return parsed.isOk()
        ? Result.ok(parsed.value)
        : Result.err(usage(`--${flag.flag} expects a JSON value.`));
    }
    default: {
      const unreachable: never = flag.valueType;
      return panic("Unhandled flag value type", { unreachable });
    }
  }
};

type ReadInputOptions = {
  source: string;
  readStdin: () => Promise<string>;
};

/**
 * Read `--input`: `@path` reads a file, `-` reads stdin, anything else is
 * inline JSON. The result must be a JSON object or array.
 */
export const readInputSource = async ({
  source,
  readStdin,
}: ReadInputOptions): Promise<Result<Record<string, unknown> | unknown[], FolioCliError>> => {
  let text = source;
  if (source === "-") {
    text = await readStdin();
  } else if (source.startsWith("@")) {
    const inputPath = source.slice(1);
    const read = await Result.tryPromise({
      try: () => readFile(inputPath, "utf8"),
      catch: () => usage(`Cannot read --input file ${inputPath}.`),
    });
    if (read.isErr()) return Result.err(read.error);
    text = read.value;
  }
  if (Buffer.byteLength(text, "utf8") > MAX_INPUT_BYTES) {
    return Result.err(usage(`--input is over the ${MAX_INPUT_BYTES}-byte limit.`));
  }
  const parsed = Result.try((): unknown => JSON.parse(text));
  if (parsed.isErr()) {
    return Result.err(
      usage("--input is not valid JSON.", "Pass inline JSON, @file.json, or - for stdin."),
    );
  }
  const value = parsed.value;
  if (Array.isArray(value)) {
    const items: unknown[] = value;
    return Result.ok(items);
  }
  return isRecord(value)
    ? Result.ok(value)
    : Result.err(usage("--input must be a JSON object or array."));
};
