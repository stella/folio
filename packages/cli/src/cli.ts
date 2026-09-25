/**
 * `folio`: the command line over the file tool registry. Each registry
 * command parses its generated flags, runs the same executor the MCP server
 * uses, and prints the shared envelope.
 */

import { Result } from "better-result";
import type { Readable, Writable } from "node:stream";
import { parseArgs, type ParseArgsOptionDescriptor } from "node:util";

import packageJson from "../package.json" with { type: "json" };
import {
  cliError,
  EXIT_CODES,
  exitCodeEntries,
  exitCodeForError,
  FOLIO_CLI_ERROR_CODES,
  type FolioCliError,
} from "./errors";
import { CLI_READ_BOUNDS, executeReadTool } from "./execute-read";
import { executeWriteTool, type WriteDestination } from "./execute-write";
import { coerceFlag, flagsForSchema, readInputSource, type GeneratedFlag } from "./flags";
import { serveFolioMcp } from "./mcp";
import { runRender, runServe } from "./preview-commands";
import {
  failureEnvelope,
  isOutputFormat,
  OUTPUT_FORMATS,
  renderEnvelope,
  successEnvelope,
  type OutputFormat,
} from "./output";
import { resolveAuthor, resolveTransactionDate } from "./provenance";
import { findCommand, listCommands, toolAccess, type FolioResolvedCommand } from "./registry";
import { resolveRoots } from "./roots";

/** The process surface the command line reads from and writes to. */
export type FolioCliIo = {
  stdout: (text: string) => void;
  stderr: (text: string) => void;
  readStdin: () => Promise<string>;
  /** Whether stdout is a terminal; picks the default output format. */
  isTTY: boolean;
  /** Environment for `FOLIO_AUTHOR`. */
  env: Readonly<Record<string, string | undefined>>;
  /** Directory whose git `user.name` is the last author fallback, and the default MCP root. */
  cwd: string;
  /** Byte streams `folio mcp` speaks the protocol over; absent, it refuses to start. */
  stdio?: { input: Readable; output: Writable };
  /** Resolves when the process is asked to stop; `folio serve` runs until then. */
  untilInterrupted?: () => Promise<void>;
};

type ParsedValue = string | boolean | (string | boolean)[] | undefined;
type ParsedValues = Record<string, ParsedValue>;

type CommonFlag = { flag: string; description: string; option: ParseArgsOptionDescriptor };

const COMMON_FLAGS: readonly CommonFlag[] = [
  {
    flag: "output",
    description: `Output format: ${OUTPUT_FORMATS.join(" or ")} (default: text on a terminal, json otherwise).`,
    option: { type: "string" },
  },
  {
    flag: "input",
    description: "All arguments as a JSON object: inline, @file.json, or - for stdin.",
    option: { type: "string" },
  },
  {
    flag: "expect-version",
    description: "Refuse unless the file's SHA-256 fileVersion is this value.",
    option: { type: "string" },
  },
  {
    flag: "help",
    description: "Show help for this command.",
    option: { type: "boolean", short: "h" },
  },
];

const WRITE_FLAGS: readonly CommonFlag[] = [
  {
    flag: "in-place",
    description: "Write the change to the file itself, keeping a backup in .folio/backups.",
    option: { type: "boolean" },
  },
  {
    flag: "out",
    description:
      "Write the result to this path instead (refused if it exists, unless --overwrite).",
    option: { type: "string", short: "o" },
  },
  {
    flag: "overwrite",
    description:
      "Let -o replace an existing file; needs --expect-destination-version. The replaced file is backed up.",
    option: { type: "boolean" },
  },
  {
    flag: "expect-destination-version",
    description: "The fileVersion of the file -o --overwrite replaces; refused if it changed.",
    option: { type: "string" },
  },
  {
    flag: "no-expect-version",
    description:
      "Change the file without naming its fileVersion. A change otherwise needs --expect-version.",
    option: { type: "boolean" },
  },
  {
    flag: "author",
    description: "Author of the change (default: FOLIO_AUTHOR, then git user.name).",
    option: { type: "string" },
  },
  {
    flag: "date",
    description: "ISO-8601 timestamp for the change (default: now), for reproducible output.",
    option: { type: "string" },
  },
  {
    flag: "allow-repack",
    description: "Allow rewriting the whole package when the edit cannot be patched in.",
    option: { type: "boolean" },
  },
  {
    flag: "tx-id",
    description: "Idempotency key: re-running a committed transaction returns its receipt.",
    option: { type: "string" },
  },
  {
    flag: "force",
    description: "Take the write lease over from another live holder.",
    option: { type: "boolean" },
  },
  {
    flag: "journal",
    description: "Journal file (default: .folio/journal.jsonl beside the destination).",
    option: { type: "string" },
  },
];

const DIRECT_FLAG: CommonFlag = {
  flag: "direct",
  description: "Edit the text directly instead of as tracked changes.",
  option: { type: "boolean" },
};

const usageError = (message: string, hint?: string): FolioCliError =>
  cliError({ code: FOLIO_CLI_ERROR_CODES.usage, message, hint });

const writeFlagsFor = ({ tool }: FolioResolvedCommand): readonly CommonFlag[] => {
  if (toolAccess(tool) === "read") return [];
  return tool.type === "agentWrite" && tool.editMode === "tracked-or-direct"
    ? [...WRITE_FLAGS, DIRECT_FLAG]
    : WRITE_FLAGS;
};

const generatedFlags = ({ tool, command }: FolioResolvedCommand): GeneratedFlag[] => {
  const bound = new Set([
    ...Object.keys(command.preset ?? {}),
    ...(command.positionals ?? []).map(({ property }) => property),
  ]);
  return flagsForSchema(tool.argsSchema, command.flagNames).filter(
    ({ property }) => !bound.has(property),
  );
};

const parseOptions = (
  resolved: FolioResolvedCommand,
): Record<string, ParseArgsOptionDescriptor> => {
  const options: Record<string, ParseArgsOptionDescriptor> = {};
  for (const { flag, option } of [...COMMON_FLAGS, ...writeFlagsFor(resolved)]) {
    options[flag] = option;
  }
  for (const { flag, valueType } of generatedFlags(resolved)) {
    options[flag] =
      valueType === "boolean"
        ? { type: "boolean" }
        : { type: "string", ...(valueType === "strings" && { multiple: true }) };
  }
  return options;
};

const wrapText = (value: string, indent: string): string => value.replaceAll("\n", `\n${indent}`);

const usageLine = ({ command }: FolioResolvedCommand): string =>
  [
    "folio",
    command.name,
    "<file>",
    ...(command.positionals ?? []).map(({ name }) => `<${name}>`),
    "[flags]",
  ].join(" ");

const rootHelp = (): string => {
  const commands = listCommands().map(
    ({ command }) => `  ${command.name.padEnd(12)}${command.summary}`,
  );
  const exitCodes = exitCodeEntries().map(
    ({ code, meaning }) => `  ${String(code).padEnd(4)}${meaning}`,
  );
  return [
    `folio ${packageJson.version}: read, review, and redline .docx files.`,
    "",
    "Usage: folio <command> <file> [flags]",
    "",
    "Commands:",
    ...commands,
    `  ${"render".padEnd(12)}Render pages to PDF, PNG, or HTML`,
    `  ${"serve".padEnd(12)}Serve a read-only live preview on 127.0.0.1`,
    `  ${"mcp".padEnd(12)}Serve these tools over MCP on stdio`,
    "",
    "Every command prints { ok, data } or { ok, error: { code, message, hint } }.",
    "A change is written only with --in-place or -o <path>, as tracked changes unless --direct.",
    "Run folio <command> --help for its flags.",
    "",
    "Exit codes:",
    ...exitCodes,
    "",
  ].join("\n");
};

const commandHelp = (resolved: FolioResolvedCommand): string => {
  const flags = [
    ...generatedFlags(resolved).map(({ flag, description, valueType }) => ({
      name: valueType === "boolean" ? `--${flag}` : `--${flag} <${valueType}>`,
      description,
    })),
    ...[...writeFlagsFor(resolved), ...COMMON_FLAGS].map(({ flag, description, option }) => ({
      name: `${option.short === undefined ? "" : `-${option.short}, `}--${flag}${
        option.type === "boolean" ? "" : " <value>"
      }`,
      description,
    })),
  ];
  return [
    `Usage: ${usageLine(resolved)}`,
    "",
    resolved.command.summary,
    "",
    wrapText(resolved.tool.description, ""),
    "",
    "Flags:",
    ...flags.map(({ name, description }) => `  ${name}\n      ${wrapText(description, "      ")}`),
    "",
  ].join("\n");
};

type CommandArguments = {
  file: string;
  expectVersion: string | undefined;
  args: Record<string, unknown>;
};

type BuildArgumentsOptions = {
  resolved: FolioResolvedCommand;
  values: ParsedValues;
  positionals: readonly string[];
  io: FolioCliIo;
};

const stringValue = (value: ParsedValue): string | undefined =>
  typeof value === "string" ? value : undefined;

const flagValue = (value: ParsedValue): string | boolean | readonly string[] | undefined => {
  if (!Array.isArray(value)) return value;
  return value.filter((entry): entry is string => typeof entry === "string");
};

const buildArguments = async ({
  resolved,
  values,
  positionals,
  io,
}: BuildArgumentsOptions): Promise<Result<CommandArguments, FolioCliError>> => {
  const { command } = resolved;
  const [file, ...rest] = positionals;
  const named = command.positionals ?? [];
  if (file === undefined || rest.length < named.length) {
    return Result.err(usageError(`Usage: ${usageLine(resolved)}.`));
  }
  if (rest.length > named.length) {
    return Result.err(usageError(`Unexpected argument ${JSON.stringify(rest[named.length])}.`));
  }
  let args: Record<string, unknown> = {};
  const input = stringValue(values["input"]);
  if (input !== undefined) {
    const parsed = await readInputSource({ source: input, readStdin: io.readStdin });
    if (parsed.isErr()) return Result.err(parsed.error);
    if (Array.isArray(parsed.value)) {
      if (command.inputArrayProperty === undefined) {
        return Result.err(usageError("--input must be a JSON object for this command."));
      }
      args = { [command.inputArrayProperty]: parsed.value };
    } else {
      args = { ...parsed.value };
    }
  }
  for (const flag of generatedFlags(resolved)) {
    const raw = flagValue(values[flag.flag]);
    if (raw === undefined) continue;
    const coerced = coerceFlag({ flag, raw });
    if (coerced.isErr()) return Result.err(coerced.error);
    args[flag.property] = coerced.value;
  }
  named.forEach(({ property }, index) => {
    args[property] = rest[index];
  });
  return Result.ok({
    file,
    expectVersion: stringValue(values["expect-version"]),
    args: { ...args, ...command.preset },
  });
};

/** `--in-place` or `-o`, exactly one; `null` when neither is given. */
const destinationFrom = (values: ParsedValues): Result<WriteDestination | null, FolioCliError> => {
  const inPlace = values["in-place"] === true;
  const out = stringValue(values["out"]);
  if (inPlace && out !== undefined) {
    return Result.err(usageError("Pass --in-place or -o <path>, not both."));
  }
  const expectedVersion = stringValue(values["expect-destination-version"]);
  if ((values["overwrite"] === true || expectedVersion !== undefined) && out === undefined) {
    return Result.err(
      usageError("--overwrite and --expect-destination-version only apply to -o <path>."),
    );
  }
  if (inPlace) return Result.ok({ type: "inPlace" });
  if (out !== undefined) {
    return Result.ok({
      type: "file",
      path: out,
      overwrite: values["overwrite"] === true,
      expectedVersion,
    });
  }
  return Result.ok(null);
};

const defaultFormat = (io: FolioCliIo): OutputFormat => (io.isTTY ? "text" : "json");

/** The `--output` value named in raw argv, so a parse failure still prints in it. */
const requestedFormat = (argv: readonly string[], io: FolioCliIo): OutputFormat => {
  const index = argv.indexOf("--output");
  const inline = argv.find((arg) => arg.startsWith("--output="))?.slice("--output=".length);
  const value = inline ?? (index === -1 ? undefined : argv[index + 1]);
  return isOutputFormat(value) ? value : defaultFormat(io);
};

type EmitOptions = {
  io: FolioCliIo;
  format: OutputFormat;
  tool: string;
  result: Result<unknown, FolioCliError>;
};

const emit = ({ io, format, tool, result }: EmitOptions): number => {
  const envelope = result.isOk() ? successEnvelope(result.value) : failureEnvelope(result.error);
  const rendered = renderEnvelope({ envelope, format, tool });
  if (rendered.stdout !== "") io.stdout(rendered.stdout);
  if (rendered.stderr !== "") io.stderr(rendered.stderr);
  return result.isOk() ? EXIT_CODES.ok : exitCodeForError(result.error.code);
};

type ExecuteOptions = {
  resolved: FolioResolvedCommand;
  values: ParsedValues;
  built: CommandArguments;
  io: FolioCliIo;
};

const execute = async ({
  resolved,
  values,
  built,
  io,
}: ExecuteOptions): Promise<Result<unknown, FolioCliError>> => {
  const { tool, command } = resolved;
  const call = { path: built.file, fileVersion: built.expectVersion, args: built.args };
  const access = toolAccess(tool);
  if (access === "read") {
    return await executeReadTool(tool, call, CLI_READ_BOUNDS);
  }
  const destination = destinationFrom(values);
  if (destination.isErr()) return destination;
  if (destination.value === null) {
    if (access === "readOrWrite") {
      return await executeReadTool(tool, call, CLI_READ_BOUNDS);
    }
    return Result.err(
      usageError(
        `folio ${command.name} changes the document: pass --in-place or -o <path>.`,
        "--in-place keeps a backup in .folio/backups beside the file.",
      ),
    );
  }
  const author = resolveAuthor({
    explicit: stringValue(values["author"]),
    env: io.env,
    cwd: io.cwd,
  });
  if (author.isErr()) return author;
  const date = resolveTransactionDate(stringValue(values["date"]));
  if (date.isErr()) return date;
  const waived = values["no-expect-version"] === true;
  if (waived && built.expectVersion !== undefined) {
    return Result.err(usageError("Pass --expect-version or --no-expect-version, not both."));
  }
  return await executeWriteTool(tool, call, {
    sourcePrecondition: waived ? "waived" : "required",
    destination: destination.value,
    author: author.value,
    date: date.value,
    repack: values["allow-repack"] === true ? "allow" : "refuse",
    force: values["force"] === true,
    txId: stringValue(values["tx-id"]),
    journalPath: stringValue(values["journal"]),
    mode: values["direct"] === true ? "direct" : "tracked-changes",
  });
};

const runCommand = async (
  resolved: FolioResolvedCommand,
  rest: readonly string[],
  io: FolioCliIo,
): Promise<number> => {
  const fallbackFormat = requestedFormat(rest, io);
  const parsed = Result.try({
    try: () =>
      parseArgs({
        args: [...rest],
        options: parseOptions(resolved),
        allowPositionals: true,
        strict: true,
      }),
    catch: (error) =>
      usageError(
        error instanceof Error ? error.message : String(error),
        `Run folio ${resolved.command.name} --help.`,
      ),
  });
  if (parsed.isErr()) {
    return emit({ io, format: fallbackFormat, tool: resolved.tool.name, result: parsed });
  }
  const { values, positionals } = parsed.value;
  if (values["help"] === true) {
    io.stdout(commandHelp(resolved));
    return EXIT_CODES.ok;
  }
  const output = values["output"];
  if (output !== undefined && !isOutputFormat(output)) {
    return emit({
      io,
      format: defaultFormat(io),
      tool: resolved.tool.name,
      result: Result.err(usageError(`--output must be ${OUTPUT_FORMATS.join(" or ")}.`)),
    });
  }
  const format = output ?? defaultFormat(io);
  const built = await buildArguments({ resolved, values, positionals, io });
  if (built.isErr()) {
    return emit({ io, format, tool: resolved.tool.name, result: built });
  }
  const result = await execute({ resolved, values, built: built.value, io });
  return emit({ io, format, tool: resolved.tool.name, result });
};

const MCP_HELP = [
  "Usage: folio mcp [--root <dir>]... [--author <name>]",
  "",
  "Serve the folio tools over MCP on stdin/stdout. Diagnostics go to stderr.",
  "",
  "Flags:",
  "  --root <dir>",
  "      A directory tool calls may read and write under (repeatable; default: the current directory).",
  "  --author <name>",
  "      Author of every change (default: FOLIO_AUTHOR, then git user.name). Without one, changes are refused.",
  "  -h, --help",
  "      Show this help.",
  "",
].join("\n");

const runMcp = async (rest: readonly string[], io: FolioCliIo): Promise<number> => {
  const fail = (error: FolioCliError): number => {
    const hint = error.hint === undefined ? "" : `hint: ${error.hint}\n`;
    io.stderr(`error: ${error.message}\n${hint}`);
    return exitCodeForError(error.code);
  };
  const parsed = Result.try({
    try: () =>
      parseArgs({
        args: [...rest],
        options: {
          root: { type: "string", multiple: true },
          author: { type: "string" },
          help: { type: "boolean", short: "h" },
        },
        strict: true,
      }),
    catch: (error) => usageError(error instanceof Error ? error.message : String(error)),
  });
  if (parsed.isErr()) return fail(parsed.error);
  const { values } = parsed.value;
  if (values.help === true) {
    io.stdout(MCP_HELP);
    return EXIT_CODES.ok;
  }
  const roots = await resolveRoots(values.root ?? [io.cwd]);
  if (roots.isErr()) return fail(roots.error);
  if (io.stdio === undefined) {
    return fail(usageError("folio mcp needs stdin and stdout streams."));
  }
  const author = resolveAuthor({ explicit: values.author, env: io.env, cwd: io.cwd });
  if (author.isErr()) {
    io.stderr(`folio mcp: ${author.error.message} Changes will be refused.\n`);
  }
  io.stderr(`folio mcp ${packageJson.version}: roots ${roots.value.join(", ")}\n`);
  await serveFolioMcp({
    input: io.stdio.input,
    output: io.stdio.output,
    roots: roots.value,
    author: author.isOk() ? author.value : undefined,
  });
  return EXIT_CODES.ok;
};

/** Run `folio` with the given arguments (without the executable); resolves to the exit code. */
export const runFolioCli = async (argv: readonly string[], io: FolioCliIo): Promise<number> => {
  const [name, ...rest] = argv;
  if (name === undefined || name === "help" || name === "--help" || name === "-h") {
    io.stdout(rootHelp());
    return EXIT_CODES.ok;
  }
  if (name === "--version" || name === "-V") {
    io.stdout(`${packageJson.version}\n`);
    return EXIT_CODES.ok;
  }
  if (name === "mcp") {
    return await runMcp(rest, io);
  }
  if (name === "render") {
    return await runRender(rest, io);
  }
  if (name === "serve") {
    return await runServe(rest, io);
  }
  const resolved = findCommand(name);
  if (resolved === undefined) {
    return emit({
      io,
      format: requestedFormat(rest, io),
      tool: name,
      result: Result.err(
        usageError(`Unknown command ${JSON.stringify(name)}.`, "Run folio --help."),
      ),
    });
  }
  return await runCommand(resolved, rest, io);
};
