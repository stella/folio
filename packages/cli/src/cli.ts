/**
 * `folio`: the command line over the file tool registry. Each registry
 * command parses its generated flags, runs the same executor the MCP server
 * uses, and prints the shared envelope.
 */

import { Result } from "better-result";
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
import { coerceFlag, flagsForSchema, readInputSource, type GeneratedFlag } from "./flags";
import {
  failureEnvelope,
  isOutputFormat,
  OUTPUT_FORMATS,
  renderEnvelope,
  successEnvelope,
  type OutputFormat,
} from "./output";
import { findCommand, listCommands, type FolioResolvedCommand } from "./registry";

/** The process surface the command line reads from and writes to. */
export type FolioCliIo = {
  stdout: (text: string) => void;
  stderr: (text: string) => void;
  readStdin: () => Promise<string>;
  /** Whether stdout is a terminal; picks the default output format. */
  isTTY: boolean;
};

type ParsedValues = Record<string, string | boolean | undefined>;

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

const usageError = (message: string, hint?: string): FolioCliError =>
  cliError({ code: FOLIO_CLI_ERROR_CODES.usage, message, hint });

const generatedFlags = ({ tool, command }: FolioResolvedCommand): GeneratedFlag[] =>
  flagsForSchema(tool.argsSchema).filter(
    ({ property }) => command.preset === undefined || !(property in command.preset),
  );

const parseOptions = (
  resolved: FolioResolvedCommand,
): Record<string, ParseArgsOptionDescriptor> => {
  const options: Record<string, ParseArgsOptionDescriptor> = {};
  for (const { flag, option } of COMMON_FLAGS) {
    options[flag] = option;
  }
  for (const { flag, valueType } of generatedFlags(resolved)) {
    options[flag] = { type: valueType === "boolean" ? "boolean" : "string" };
  }
  return options;
};

const wrapText = (value: string, indent: string): string => value.replaceAll("\n", `\n${indent}`);

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
    "",
    "Every command prints { ok, data } or { ok, error: { code, message, hint } }.",
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
    ...COMMON_FLAGS.map(({ flag, description, option }) => ({
      name: option.type === "boolean" ? `--${flag}` : `--${flag} <value>`,
      description,
    })),
  ];
  return [
    `Usage: folio ${resolved.command.name} <file> [flags]`,
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

const buildArguments = async ({
  resolved,
  values,
  positionals,
  io,
}: BuildArgumentsOptions): Promise<Result<CommandArguments, FolioCliError>> => {
  const [file, ...extra] = positionals;
  if (file === undefined) {
    return Result.err(usageError(`folio ${resolved.command.name} needs a <file>.`));
  }
  if (extra.length > 0) {
    return Result.err(usageError(`Unexpected argument ${JSON.stringify(extra[0])}.`));
  }
  let args: Record<string, unknown> = {};
  const input = values["input"];
  if (typeof input === "string") {
    const parsed = await readInputSource({ source: input, readStdin: io.readStdin });
    if (parsed.isErr()) return Result.err(parsed.error);
    if (Array.isArray(parsed.value)) {
      return Result.err(usageError("--input must be a JSON object for this command."));
    }
    args = { ...parsed.value };
  }
  for (const flag of generatedFlags(resolved)) {
    const raw = values[flag.flag];
    if (raw === undefined) continue;
    const coerced = coerceFlag({ flag, raw });
    if (coerced.isErr()) return Result.err(coerced.error);
    args[flag.property] = coerced.value;
  }
  const expectVersion = values["expect-version"];
  return Result.ok({
    file,
    expectVersion: typeof expectVersion === "string" ? expectVersion : undefined,
    args: { ...args, ...resolved.command.preset },
  });
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
  const { file, expectVersion, args } = built.value;
  const result = await executeReadTool(
    resolved.tool,
    { path: file, fileVersion: expectVersion, args },
    CLI_READ_BOUNDS,
  );
  return emit({ io, format, tool: resolved.tool.name, result });
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
