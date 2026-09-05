/**
 * Compare two `.docx` files from a terminal.
 *
 *   bun run scripts/compare.ts <base.docx> <target.docx> <out.docx> [--json]
 *
 * Writes `out.docx` — the base carrying the tracked changes that turn it into
 * the target — and prints the change list. `--json` prints it as JSON for a
 * script to consume; without it the output is one line per change.
 *
 * `--author` and `--timestamp` pin what the generated revisions are stamped
 * with. Both default to fixed values rather than the current user and clock, so
 * running the script twice over the same pair produces the same file.
 */

import { readFileSync, writeFileSync } from "node:fs";

import { compareDocx } from "../src/compare/compare";
import type { CompareChange, CompareUnsupportedPart } from "../src/compare/types";

const DEFAULT_AUTHOR = "folio compare";
const DEFAULT_TIMESTAMP = "1970-01-01T00:00:00.000Z";

const USAGE =
  "usage: bun run scripts/compare.ts <base.docx> <target.docx> <out.docx> [--json] [--author <name>] [--timestamp <iso8601>]";

type ParsedArgs = {
  basePath: string;
  targetPath: string;
  outPath: string;
  json: boolean;
  author: string;
  timestamp: string;
};

const parseArgs = (argv: readonly string[]): ParsedArgs | null => {
  const positional: string[] = [];
  let json = false;
  let author = DEFAULT_AUTHOR;
  let timestamp = DEFAULT_TIMESTAMP;

  for (let index = 0; index < argv.length; index++) {
    const arg = argv[index];
    if (arg === "--json") {
      json = true;
      continue;
    }
    if (arg === "--author" || arg === "--timestamp") {
      const value = argv[++index];
      if (value === undefined) {
        return null;
      }
      if (arg === "--author") {
        author = value;
      } else {
        timestamp = value;
      }
      continue;
    }
    if (arg === undefined || arg.startsWith("--")) {
      return null;
    }
    positional.push(arg);
  }

  const [basePath, targetPath, outPath] = positional;
  if (positional.length !== 3 || !basePath || !targetPath || !outPath) {
    return null;
  }
  return { basePath, targetPath, outPath, json, author, timestamp };
};

const readDocx = (filePath: string): ArrayBuffer => {
  const bytes = readFileSync(filePath);
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
};

const describeChange = (change: CompareChange): string => {
  const where =
    change.location.cell === undefined
      ? "body"
      : `table ${String(change.location.cell.tableIndex)} row ${String(change.location.cell.rowIndex)} cell ${String(change.location.cell.cellIndex)}`;
  switch (change.kind) {
    case "insert":
      return `insert  [${where}] ${change.after}`;
    case "delete":
      return `delete  [${where}] ${change.before}`;
    case "replace":
      return `replace [${where}] ${change.before} -> ${change.after}`;
    case "move":
      return `move    [${where}] ${change.text}`;
    case "format":
      return `format  [${where}] ${String(change.ranges.length)} range(s) in ${change.text}`;
    case "table-row-insert":
      return `row +   [table ${String(change.tableIndex)} row ${String(change.rowIndex)}] ${change.cells.join(" | ")}`;
    case "table-row-delete":
      return `row -   [table ${String(change.tableIndex)} row ${String(change.rowIndex)}] ${change.cells.join(" | ")}`;
  }
};

const describeUnsupported = ({ reason, baseStory, targetStory }: CompareUnsupportedPart): string =>
  `unsupported (${reason}): ${JSON.stringify(baseStory ?? targetStory)}`;

const describeError = (value: unknown): string => {
  if (value instanceof Error) {
    return `${value.name}: ${value.message}`;
  }
  return typeof value === "string" ? value : JSON.stringify(value);
};

/** Guard against a chain that loops or is pathologically deep. */
const MAX_CAUSE_DEPTH = 8;

const readCause = (value: unknown): unknown =>
  typeof value === "object" && value !== null && "cause" in value ? value.cause : undefined;

/**
 * The error and every cause behind it, outermost first.
 *
 * A parse failure says only that the document could not be parsed; which part
 * refused it lives in the cause. Printing the tag alone left a failing file
 * undiagnosable from the CLI.
 */
const causeChain = (error: unknown): string[] => {
  const lines: string[] = [];
  const seen = new Set<unknown>();
  let current: unknown = error;
  while (current !== undefined && current !== null && !seen.has(current)) {
    if (lines.length >= MAX_CAUSE_DEPTH) {
      lines.push("... further causes omitted");
      break;
    }
    seen.add(current);
    lines.push(describeError(current));
    current = readCause(current);
  }
  return lines;
};

const args = parseArgs(process.argv.slice(2));
if (!args) {
  console.error(USAGE);
  process.exit(2);
}

const result = await compareDocx(readDocx(args.basePath), readDocx(args.targetPath), {
  author: args.author,
  timestamp: args.timestamp,
});

if (result.isErr()) {
  console.error(`${result.error._tag}: ${result.error.message}`);
  for (const [index, line] of causeChain(readCause(result.error)).entries()) {
    console.error(`${"  ".repeat(index + 1)}caused by: ${line}`);
  }
  process.exit(1);
}

const { buffer, changes, unsupported } = result.value;
writeFileSync(args.outPath, new Uint8Array(buffer));

if (args.json) {
  console.log(JSON.stringify({ changes, unsupported }, null, 2));
} else {
  for (const change of changes) {
    console.log(describeChange(change));
  }
  for (const part of unsupported) {
    console.log(describeUnsupported(part));
  }
  console.log(`${String(changes.length)} change(s) written to ${args.outPath}`);
}
