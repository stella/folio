/**
 * `folio render` and `folio serve`: the document as its pages, from folio's
 * own layout and painters. Neither changes the document.
 */

import { Result } from "better-result";
import path from "node:path";
import { parseArgs } from "node:util";

import type { FolioCliIo } from "./cli";
import { checkExpectedVersion, readDocumentFile } from "./document";
import {
  cliError,
  EXIT_CODES,
  exitCodeForError,
  FOLIO_CLI_ERROR_CODES,
  type FolioCliError,
} from "./errors";
import {
  failureEnvelope,
  isOutputFormat,
  OUTPUT_FORMATS,
  renderEnvelope,
  successEnvelope,
  type OutputFormat,
} from "./output";
import { writeOutputFile } from "./output-file";
import { resolveTransactionDate } from "./provenance";
import {
  buildDisplayList,
  checkPages,
  displayListHtml,
  renderPdf,
  renderPng,
  type RenderFormat,
} from "./render";
import { startPreviewServer } from "./serve";

const usageError = (message: string, hint?: string): FolioCliError =>
  cliError({ code: FOLIO_CLI_ERROR_CODES.usage, message, hint });

const emit = (
  io: FolioCliIo,
  format: OutputFormat,
  result: Result<unknown, FolioCliError>,
): number => {
  const envelope = result.isOk() ? successEnvelope(result.value) : failureEnvelope(result.error);
  const rendered = renderEnvelope({ envelope, format, tool: "render" });
  if (rendered.stdout !== "") io.stdout(rendered.stdout);
  if (rendered.stderr !== "") io.stderr(rendered.stderr);
  return result.isOk() ? EXIT_CODES.ok : exitCodeForError(result.error.code);
};

const formatFrom = (value: unknown, io: FolioCliIo): Result<OutputFormat, FolioCliError> => {
  if (value === undefined) return Result.ok(io.isTTY ? "text" : "json");
  return isOutputFormat(value)
    ? Result.ok(value)
    : Result.err(usageError(`--output must be ${OUTPUT_FORMATS.join(" or ")}.`));
};

/** The output format an extension names, or `null` for any other. */
const renderFormatOf = (target: string): RenderFormat | null => {
  switch (path.extname(target).toLowerCase()) {
    case ".pdf":
      return "pdf";
    case ".png":
      return "png";
    case ".html":
      return "html";
    default:
      return null;
  }
};

const positiveInteger = (
  value: string | undefined,
  flag: string,
): Result<number | undefined, FolioCliError> => {
  if (value === undefined) return Result.ok(undefined);
  const parsed = Number(value);
  return Number.isInteger(parsed) && parsed >= 1
    ? Result.ok(parsed)
    : Result.err(usageError(`${flag} expects a whole number from 1.`));
};

export const RENDER_HELP = [
  "Usage: folio render <file> -o <out.pdf|out.png|out.html> [--page <n>]",
  "",
  "Render the document with folio's layout: a PDF of every page (or --page n), an HTML",
  "document of its pages, or one page as a PNG (page 1 unless --page). PNG needs Chromium",
  "through playwright-core. The document is never changed.",
  "",
  "Flags:",
  "  -o, --out <path>          Output file; its extension picks the format",
  "  --page <n>                Only this page (1-based)",
  "  --scale <n>               PNG pixels per CSS pixel (default 2)",
  "  --overwrite               Replace an existing output file",
  "  --expect-version <sha>    Refuse unless the document has this fileVersion",
  "  --date <iso>              PDF creation date (default: now)",
  "  --output <json|text>      Envelope format",
  "",
].join("\n");

const renderOptions = {
  out: { type: "string", short: "o" },
  page: { type: "string" },
  scale: { type: "string" },
  overwrite: { type: "boolean" },
  "expect-version": { type: "string" },
  date: { type: "string" },
  output: { type: "string" },
  help: { type: "boolean", short: "h" },
} as const;

export const runRender = async (rest: readonly string[], io: FolioCliIo): Promise<number> => {
  const parsed = Result.try({
    try: () =>
      parseArgs({ args: [...rest], options: renderOptions, allowPositionals: true, strict: true }),
    catch: (error) =>
      usageError(
        error instanceof Error ? error.message : String(error),
        "Run folio render --help.",
      ),
  });
  if (parsed.isErr()) return emit(io, io.isTTY ? "text" : "json", parsed);
  const { values, positionals } = parsed.value;
  if (values.help === true) {
    io.stdout(RENDER_HELP);
    return EXIT_CODES.ok;
  }
  const format = formatFrom(values.output, io);
  if (format.isErr()) return emit(io, io.isTTY ? "text" : "json", format);
  const result = await render({ values, positionals });
  return emit(io, format.value, result);
};

type RenderArguments = {
  values: {
    out?: string | undefined;
    page?: string | undefined;
    scale?: string | undefined;
    overwrite?: boolean | undefined;
    "expect-version"?: string | undefined;
    date?: string | undefined;
  };
  positionals: readonly string[];
};

const render = async ({
  values,
  positionals,
}: RenderArguments): Promise<Result<unknown, FolioCliError>> => {
  const [input, ...extra] = positionals;
  if (input === undefined || extra.length > 0) {
    return Result.err(usageError("Usage: folio render <file> -o <out.pdf|out.png|out.html>."));
  }
  if (values.out === undefined) {
    return Result.err(usageError("folio render needs -o <out.pdf|out.png|out.html>."));
  }
  const kind = renderFormatOf(values.out);
  if (kind === null) {
    return Result.err(usageError("-o must end in .pdf, .png, or .html."));
  }
  const page = positiveInteger(values.page, "--page");
  if (page.isErr()) return Result.err(page.error);
  const scale = positiveInteger(values.scale, "--scale");
  if (scale.isErr()) return Result.err(scale.error);
  const date = resolveTransactionDate(values.date);
  if (date.isErr()) return Result.err(date.error);

  const file = await readDocumentFile(input);
  if (file.isErr()) return Result.err(file.error);
  const version = checkExpectedVersion(file.value, values["expect-version"]);
  if (version.isErr()) return Result.err(version.error);

  const pages = page.value === undefined ? undefined : [page.value];
  let output: Result<{ bytes: Uint8Array; pageCount: number }, FolioCliError>;
  switch (kind) {
    case "pdf":
      output = await renderPdf({ file: file.value, pages, timestamp: date.value });
      break;
    case "png":
      output = await renderPng({
        file: file.value,
        page: page.value ?? 1,
        scale: scale.value ?? 2,
      });
      break;
    case "html": {
      const list = await buildDisplayList(file.value);
      const selected = list.isOk() ? checkPages(list.value, pages) : list;
      output = selected.isOk()
        ? Result.ok({
            bytes: new TextEncoder().encode(
              displayListHtml(selected.value, path.basename(file.value.path)),
            ),
            pageCount: selected.value.pages.length,
          })
        : Result.err(selected.error);
      break;
    }
    default: {
      const unreachable: never = kind;
      return unreachable;
    }
  }
  if (output.isErr()) return Result.err(output.error);
  const written = await writeOutputFile({
    target: values.out,
    bytes: output.value.bytes,
    overwrite: values.overwrite === true,
    input: file.value.identity,
  });
  if (written.isErr()) return Result.err(written.error);
  return Result.ok({
    path: file.value.path,
    fileVersion: file.value.fileVersion,
    output: written.value,
    format: kind,
    pageCount: output.value.pageCount,
    bytes: output.value.bytes.byteLength,
  });
};

export const SERVE_HELP = [
  "Usage: folio serve <file> [--port <n>]",
  "",
  "Serve a read-only live preview of the document on 127.0.0.1. The page follows the file:",
  "when its version changes (a folio write, or any other program), the preview re-renders.",
  "The URL carries a random token; the server never writes. Stop it with Ctrl-C.",
  "",
  "Flags:",
  "  --port <n>              Port on 127.0.0.1 (default: a free one)",
  "  --output <json|text>    Envelope format for the startup line",
  "",
].join("\n");

export const runServe = async (rest: readonly string[], io: FolioCliIo): Promise<number> => {
  const parsed = Result.try({
    try: () =>
      parseArgs({
        args: [...rest],
        options: {
          port: { type: "string" },
          output: { type: "string" },
          help: { type: "boolean", short: "h" },
        },
        allowPositionals: true,
        strict: true,
      }),
    catch: (error) =>
      usageError(error instanceof Error ? error.message : String(error), "Run folio serve --help."),
  });
  if (parsed.isErr()) return emit(io, io.isTTY ? "text" : "json", parsed);
  const { values, positionals } = parsed.value;
  if (values.help === true) {
    io.stdout(SERVE_HELP);
    return EXIT_CODES.ok;
  }
  const format = formatFrom(values.output, io);
  if (format.isErr()) return emit(io, io.isTTY ? "text" : "json", format);
  const [input, ...extra] = positionals;
  if (input === undefined || extra.length > 0) {
    return emit(
      io,
      format.value,
      Result.err(usageError("Usage: folio serve <file> [--port <n>].")),
    );
  }
  const port = values.port === undefined ? 0 : Number(values.port);
  if (!Number.isInteger(port) || port < 0 || port > 65_535) {
    return emit(io, format.value, Result.err(usageError("--port expects a port number.")));
  }
  if (io.untilInterrupted === undefined) {
    return emit(
      io,
      format.value,
      Result.err(usageError("folio serve needs to run until interrupted.")),
    );
  }
  const started = await startPreviewServer({ documentPath: input, port });
  if (started.isErr()) {
    return emit(
      io,
      format.value,
      Result.err(
        cliError({ code: FOLIO_CLI_ERROR_CODES.invalidInput, message: started.error.message }),
      ),
    );
  }
  const exit = emit(
    io,
    format.value,
    Result.ok({ url: started.value.url, path: path.resolve(input) }),
  );
  await io.untilInterrupted();
  await started.value.close();
  return exit;
};
