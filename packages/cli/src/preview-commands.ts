import { panic, Result } from "better-result";
/**
 * `folio render` and `folio serve`: the document as its pages, from folio's
 * own layout and painters. Neither changes the document.
 */
import path from "node:path";
import { parseArgs } from "node:util";

import { encodeUtf8 } from "@stll/folio-core/utils/textEncoding";

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
  type OutputFormat,
  renderEnvelope,
  successEnvelope,
} from "./output";
import { writeOutputFile } from "./output-file";
import { resolveTransactionDate } from "./provenance";
import {
  buildDisplayList,
  checkPages,
  displayListHtml,
  RENDER_FORMATS,
  type RenderFormat,
  renderPdf,
  renderPng,
} from "./render";
import { startPreviewServer } from "./serve";

const usageError = (message: string, hint?: string): FolioCliError =>
  cliError({ code: FOLIO_CLI_ERROR_CODES.usage, message, hint });

const emit = (
  io: FolioCliIo,
  format: OutputFormat,
  result: Result<unknown, FolioCliError>,
): number => {
  const envelope = result.isOk()
    ? successEnvelope(result.value)
    : failureEnvelope(result.error);
  const rendered = renderEnvelope({ envelope, format, tool: "render" });
  if (rendered.stdout !== "") io.stdout(rendered.stdout);
  if (rendered.stderr !== "") io.stderr(rendered.stderr);
  return result.isOk() ? EXIT_CODES.ok : exitCodeForError(result.error.code);
};

const formatFrom = (
  value: unknown,
  io: FolioCliIo,
): Result<OutputFormat, FolioCliError> => {
  if (value === undefined) return Result.ok(io.isTTY ? "text" : "json");
  return isOutputFormat(value)
    ? Result.ok(value)
    : Result.err(
        usageError(`--output must be ${OUTPUT_FORMATS.join(" or ")}.`),
      );
};

const isRenderFormat = (value: unknown): value is RenderFormat =>
  RENDER_FORMATS.some((format) => format === value);

/** The output format an extension names, or `null` for any other. */
const renderFormatOf = (target: string): RenderFormat | null => {
  const extension = path.extname(target).toLowerCase().slice(1);
  return isRenderFormat(extension) ? extension : null;
};

/** The placeholder in `-o` that numbers one output file per page. */
const PAGE_PLACEHOLDER = "{page}";

/**
 * One-based page numbers from a list of numbers and ranges (`1-3,5`), in
 * the order given, each once.
 */
const pagesFrom = (spec: string): Result<readonly number[], FolioCliError> => {
  const pages: number[] = [];
  for (const part of spec.split(",")) {
    const range = /^\s*(\d+)\s*(?:-\s*(\d+)\s*)?$/u.exec(part);
    const first = Number(range?.[1]);
    const last = range?.[2] === undefined ? first : Number(range[2]);
    if (range === null || first < 1 || last < first) {
      return Result.err(
        usageError(
          `--pages expects page numbers and ranges such as 1-3,5; got ${JSON.stringify(spec)}.`,
        ),
      );
    }
    for (let page = first; page <= last; page += 1) {
      if (!pages.includes(page)) pages.push(page);
    }
  }
  return Result.ok(pages);
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

export const RENDER_HELP: string = [
  "Usage: folio render <file> -o <out.pdf|out.png|out.html> [--pages <list>] [--format <pdf|png|html>]",
  "",
  "Render the document with folio's layout: a PDF of the pages, an HTML document of them, or",
  "PNG images (page 1 unless --pages; one file per page when -o contains {page}). PNG needs",
  "Chromium through playwright-core. The document is never changed.",
  "",
  "Flags:",
  "  -o, --out <path>          Output file; with {page}, one PNG per page (page-{page}.png)",
  "  --format <pdf|png|html>   Output format (default: from the -o extension)",
  "  --pages <list>            Only these pages: numbers and ranges, 1-based (1-3,5)",
  "  --page <n>                Only this page (same as --pages n)",
  "  --scale <n>               PNG pixels per CSS pixel (default 2)",
  "  --overwrite               Replace existing output files",
  "  --expect-version <sha>    Refuse unless the document has this fileVersion",
  "  --date <iso>              PDF creation date (default: now)",
  "  --output <json|text>      Envelope format",
  "",
].join("\n");

const renderOptions = {
  out: { type: "string", short: "o" },
  format: { type: "string" },
  pages: { type: "string" },
  page: { type: "string" },
  scale: { type: "string" },
  overwrite: { type: "boolean" },
  "expect-version": { type: "string" },
  date: { type: "string" },
  output: { type: "string" },
  help: { type: "boolean", short: "h" },
} as const;

export const runRender = async (
  rest: readonly string[],
  io: FolioCliIo,
): Promise<number> => {
  const parsed = Result.try({
    try: () =>
      parseArgs({
        args: [...rest],
        options: renderOptions,
        allowPositionals: true,
        strict: true,
      }),
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
    format?: string | undefined;
    pages?: string | undefined;
    page?: string | undefined;
    scale?: string | undefined;
    overwrite?: boolean | undefined;
    "expect-version"?: string | undefined;
    date?: string | undefined;
  };
  positionals: readonly string[];
};

/** The format `--format` names, or the one `-o`'s extension names; they must agree. */
const outputFormatFrom = (
  out: string,
  format: string | undefined,
): Result<RenderFormat, FolioCliError> => {
  const named = renderFormatOf(out.replaceAll(PAGE_PLACEHOLDER, ""));
  if (format === undefined) {
    return named === null
      ? Result.err(
          usageError(
            "-o must end in .pdf, .png, or .html, or --format must name the format.",
          ),
        )
      : Result.ok(named);
  }
  if (!isRenderFormat(format)) {
    return Result.err(
      usageError(`--format must be ${RENDER_FORMATS.join(", ")}.`),
    );
  }
  return named !== null && named !== format
    ? Result.err(usageError(`-o ends in .${named} but --format is ${format}.`))
    : Result.ok(format);
};

/** The pages asked for by --pages or --page; `undefined` for every page. */
const selectedPages = (
  values: RenderArguments["values"],
): Result<readonly number[] | undefined, FolioCliError> => {
  if (values.pages !== undefined && values.page !== undefined) {
    return Result.err(usageError("Pass --pages or --page, not both."));
  }
  if (values.pages !== undefined) return pagesFrom(values.pages);
  const page = positiveInteger(values.page, "--page");
  if (page.isErr()) return Result.err(page.error);
  return Result.ok(page.value === undefined ? undefined : [page.value]);
};

type RenderedFile = { page?: number; bytes: Uint8Array; target: string };

const render = async ({
  values,
  positionals,
}: RenderArguments): Promise<Result<unknown, FolioCliError>> => {
  const [input, ...extra] = positionals;
  if (input === undefined || extra.length > 0) {
    return Result.err(
      usageError("Usage: folio render <file> -o <out.pdf|out.png|out.html>."),
    );
  }
  if (values.out === undefined) {
    return Result.err(
      usageError("folio render needs -o <out.pdf|out.png|out.html>."),
    );
  }
  const out = values.out;
  const kind = outputFormatFrom(out, values.format);
  if (kind.isErr()) return Result.err(kind.error);
  const perPage = out.includes(PAGE_PLACEHOLDER);
  if (perPage && kind.value !== "png") {
    return Result.err(
      usageError(
        `${PAGE_PLACEHOLDER} in -o numbers PNG files; a ${kind.value} holds every page.`,
      ),
    );
  }
  const pages = selectedPages(values);
  if (pages.isErr()) return Result.err(pages.error);
  if (kind.value === "png" && !perPage && (pages.value?.length ?? 1) > 1) {
    return Result.err(
      usageError(
        `Several pages make several PNG files: put ${PAGE_PLACEHOLDER} in -o.`,
        "For example -o page-{page}.png.",
      ),
    );
  }
  const scale = positiveInteger(values.scale, "--scale");
  if (scale.isErr()) return Result.err(scale.error);
  const date = resolveTransactionDate(values.date);
  if (date.isErr()) return Result.err(date.error);

  const file = await readDocumentFile(input);
  if (file.isErr()) return Result.err(file.error);
  const version = checkExpectedVersion(file.value, values["expect-version"]);
  if (version.isErr()) return Result.err(version.error);

  let rendered: RenderedFile[];
  let pageCount: number;
  switch (kind.value) {
    case "pdf": {
      const pdf = await renderPdf({
        file: file.value,
        pages: pages.value,
        timestamp: date.value,
      });
      if (pdf.isErr()) return Result.err(pdf.error);
      pageCount = pdf.value.pageCount;
      rendered = [{ bytes: pdf.value.bytes, target: out }];
      break;
    }
    case "png": {
      const png = await renderPng({
        file: file.value,
        pages: perPage ? pages.value : (pages.value ?? [1]),
        scale: scale.value ?? 2,
      });
      if (png.isErr()) return Result.err(png.error);
      pageCount = png.value.length;
      rendered = png.value.map(({ page, bytes }) => ({
        page,
        bytes,
        target: out.replaceAll(PAGE_PLACEHOLDER, String(page)),
      }));
      break;
    }
    case "html": {
      const list = await buildDisplayList(file.value);
      const selected = list.isOk() ? checkPages(list.value, pages.value) : list;
      if (selected.isErr()) return Result.err(selected.error);
      pageCount = selected.value.pages.length;
      rendered = [
        {
          bytes: encodeUtf8(
            displayListHtml(selected.value, path.basename(file.value.path)),
          ),
          target: out,
        },
      ];
      break;
    }
    default: {
      const unreachable: never = kind.value;
      return panic("Unhandled render format", { unreachable });
    }
  }
  const outputs: { page?: number; output: string; bytes: number }[] = [];
  for (const { page, bytes, target } of rendered) {
    const written = await writeOutputFile({
      target,
      bytes,
      overwrite: values.overwrite === true,
      input: file.value.identity,
    });
    if (written.isErr()) {
      return Result.err(
        outputs.length === 0
          ? written.error
          : cliError({
              code: written.error.code,
              message: written.error.message,
              ...(written.error.hint !== undefined && {
                hint: written.error.hint,
              }),
              details: { written: outputs },
            }),
      );
    }
    outputs.push({
      ...(page !== undefined && { page }),
      output: written.value,
      bytes: bytes.byteLength,
    });
  }
  const [first] = outputs;
  return Result.ok({
    path: file.value.path,
    fileVersion: file.value.fileVersion,
    format: kind.value,
    pageCount,
    ...(perPage || first === undefined
      ? { outputs }
      : { output: first.output, bytes: first.bytes }),
  });
};

export const SERVE_HELP: string = [
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

export const runServe = async (
  rest: readonly string[],
  io: FolioCliIo,
): Promise<number> => {
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
      usageError(
        error instanceof Error ? error.message : String(error),
        "Run folio serve --help.",
      ),
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
    return emit(
      io,
      format.value,
      Result.err(usageError("--port expects a port number.")),
    );
  }
  if (io.untilInterrupted === undefined) {
    return emit(
      io,
      format.value,
      Result.err(usageError("folio serve needs to run until interrupted.")),
    );
  }
  const started = await startPreviewServer({ documentPath: input, port });
  if (started.isErr()) return emit(io, format.value, started);
  const exit = emit(
    io,
    format.value,
    Result.ok({ url: started.value.url, path: path.resolve(input) }),
  );
  await io.untilInterrupted();
  await started.value.close();
  return exit;
};
