/**
 * Word ground truth: export a .docx via the locally installed Microsoft Word
 * for Mac (scripted headlessly through AppleScript), extract per-line
 * geometry with `mutool draw -F stext`, and render per-page PNGs for the
 * visual report. Everything is cached under `CACHE_DIR/<sha256 of the docx
 * content>/` so a second call never reopens Word or re-runs mutool.
 */

import { mkdir, readdir, rm } from "node:fs/promises";
import path from "node:path";

import { CACHE_DIR } from "./config";
import { parseStextXml } from "./stextParse";
import { normalizeLineText } from "./textNorm";
import type { DocGeom } from "./types";

const WORD_APP_PATH = "/Applications/Microsoft Word.app";
const EXPORT_TIMEOUT_MS = 180_000;
const PNG_DPI = 96;

const PDF_FILENAME = "word.pdf";
const STEXT_XML_FILENAME = "word-stext.xml";
const GEOM_JSON_FILENAME = "word-geom.json";
const PAGES_DIRNAME = "word-pages";
const PAGE_PNG_RE = /^p(\d+)\.png$/;

const GET_WORD_VERSION_SCRIPT = 'tell application "Microsoft Word" to get version';

export type WordTruthStage = "availability" | "export" | "extract";

export class WordTruthError extends Error {
  readonly stage: WordTruthStage;

  constructor(message: string, stage: WordTruthStage) {
    super(message);
    this.name = "WordTruthError";
    this.stage = stage;
  }
}

/** Filesystem-only check: Word.app is installed AND `mutool` resolves on
 * PATH. Never launches Word. */
export const isWordAvailable = async (): Promise<boolean> => {
  if (Bun.which("mutool") === null) return false;
  return await Bun.file(path.join(WORD_APP_PATH, "Contents", "Info.plist")).exists();
};

// Cached across calls in-process: querying Word's version launches it, so we
// only pay that cost once per process.
let cachedWordVersion: string | null | undefined;

/** Word's version string via AppleScript (e.g. "16.112"), or null if Word
 * cannot be queried. Cached in a module variable after the first call. */
export const getWordVersion = async (): Promise<string | null> => {
  if (cachedWordVersion !== undefined) return cachedWordVersion;

  const proc = Bun.spawn(["osascript", "-e", GET_WORD_VERSION_SCRIPT], {
    stdout: "pipe",
    stderr: "ignore",
  });
  // Drain stdout before checking the exit code: on a non-zero exit the pipe
  // would otherwise be left unconsumed, leaking its file descriptor.
  const stdout = (await new Response(proc.stdout).text()).trim();
  const exitCode = await proc.exited;
  if (exitCode !== 0) {
    cachedWordVersion = null;
    return cachedWordVersion;
  }
  cachedWordVersion = stdout === "" ? null : stdout;
  return cachedWordVersion;
};

/** `mutool -v` prints its version to stderr, not stdout. */
const getMutoolVersion = async (): Promise<string> => {
  // `mutool -v` only writes to stderr; ignore stdout so its pipe fd is not
  // left open and leaked.
  const proc = Bun.spawn(["mutool", "-v"], { stdout: "ignore", stderr: "pipe" });
  const stderr = await new Response(proc.stderr).text();
  await proc.exited;
  return stderr.trim();
};

const sha256OfFile = async (filePath: string): Promise<string> => {
  const data = await Bun.file(filePath).arrayBuffer();
  const hasher = new Bun.CryptoHasher("sha256");
  hasher.update(data);
  return hasher.digest("hex");
};

const cacheDirFor = (sha256: string): string => path.join(CACHE_DIR, sha256);

/** Escape a path for interpolation into a double-quoted AppleScript string
 * literal: backslashes first, then double quotes. */
const escapeAppleScriptString = (value: string): string =>
  value.replace(/\\/g, "\\\\").replace(/"/g, '\\"');

// The default AppleEvent timeout (~60s) is shorter than our own process-level
// EXPORT_TIMEOUT_MS and is too short for a cold Word launch, which trips
// "AppleEvent timed out (-1712)" well before the export itself would fail.
// `with timeout of` raises the AppleEvent-level ceiling to just under our
// process timeout so a real hang is still caught (by the process kill)
// without the script's own default ceiling firing first.
const APPLE_EVENT_TIMEOUT_SECONDS = Math.floor(EXPORT_TIMEOUT_MS / 1000) - 10;

const buildExportScript = (docxPath: string, pdfPath: string): string => {
  const inFile = escapeAppleScriptString(docxPath);
  const outFile = escapeAppleScriptString(pdfPath);
  return `with timeout of ${APPLE_EVENT_TIMEOUT_SECONDS} seconds
	tell application "Microsoft Word"
		set inFile to POSIX file "${inFile}"
		open inFile
		set theDoc to active document
		save as theDoc file name "${outFile}" file format format PDF
		close theDoc saving no
	end tell
end timeout`;
};

// Word for Mac is App-Sandboxed: scripting it to open or save files in an
// arbitrary folder raises a per-location "grant access" consent dialog, which
// blocks the AppleEvent until a human clicks it (observed as -1712 timeouts).
// Files inside Word's own container are always readable/writable without any
// prompt, so both the input docx and the output PDF are staged there and the
// PDF is moved into our cache afterwards.
const WORD_CONTAINER_TMP = path.join(
  process.env["HOME"] ?? "",
  "Library",
  "Containers",
  "com.microsoft.Word",
  "Data",
  "tmp",
);

/** Export `docxPath` to `destPdfPath` by scripting Word. Both sides of the
 * conversion are staged inside Word's sandbox container (see above), and the
 * PDF is moved to `destPdfPath` only on success, so a killed/failed export
 * never leaves a half-written `word.pdf` that a later call would mistake for
 * a valid cache entry. */
const exportViaWord = async (docxPath: string, destPdfPath: string): Promise<void> => {
  const stagingToken = `parity-${process.pid}-${Date.now()}`;
  const stagedDocxPath = path.join(WORD_CONTAINER_TMP, `${stagingToken}.docx`);
  const tmpPdfPath = path.join(WORD_CONTAINER_TMP, `${stagingToken}.pdf`);
  await mkdir(WORD_CONTAINER_TMP, { recursive: true });
  await Bun.write(stagedDocxPath, Bun.file(docxPath));
  try {
    await runWordExportScript({ docxPath, stagedDocxPath, tmpPdfPath, destPdfPath });
  } finally {
    await rm(stagedDocxPath, { force: true });
  }
};

type RunWordExportArgs = {
  /** Original path, used only for error messages. */
  docxPath: string;
  stagedDocxPath: string;
  tmpPdfPath: string;
  destPdfPath: string;
};

const runWordExportScript = async (args: RunWordExportArgs): Promise<void> => {
  const { docxPath, stagedDocxPath, tmpPdfPath, destPdfPath } = args;
  const script = buildExportScript(stagedDocxPath, tmpPdfPath);

  const proc = Bun.spawn(["osascript", "-e", script], { stdout: "pipe", stderr: "pipe" });
  let timedOut = false;
  const timer = setTimeout(() => {
    timedOut = true;
    proc.kill();
  }, EXPORT_TIMEOUT_MS);

  const [exitCode, stderr] = await Promise.all([proc.exited, new Response(proc.stderr).text()]);
  clearTimeout(timer);

  if (timedOut) {
    await rm(tmpPdfPath, { force: true });
    throw new WordTruthError(
      `Word export timed out after ${EXPORT_TIMEOUT_MS}ms exporting ${docxPath}`,
      "export",
    );
  }
  if (exitCode !== 0) {
    await rm(tmpPdfPath, { force: true });
    throw new WordTruthError(
      `Word export failed for ${docxPath} (exit ${exitCode}): ${stderr.trim()}`,
      "export",
    );
  }

  const exportedFile = Bun.file(tmpPdfPath);
  if (!(await exportedFile.exists()) || exportedFile.size === 0) {
    await rm(tmpPdfPath, { force: true });
    throw new WordTruthError(`Word export produced an empty PDF for ${docxPath}`, "export");
  }

  // Cross-directory move: the container tmp dir and the cache dir may sit on
  // different volumes, so copy + remove instead of rename.
  await Bun.write(destPdfPath, exportedFile);
  await rm(tmpPdfPath, { force: true });
};

const runMutoolStext = async (pdfPath: string, xmlPath: string): Promise<void> => {
  const proc = Bun.spawn(["mutool", "draw", "-F", "stext", "-o", xmlPath, pdfPath], {
    stdout: "pipe",
    stderr: "pipe",
  });
  const [exitCode, stderr] = await Promise.all([proc.exited, new Response(proc.stderr).text()]);
  if (exitCode !== 0) {
    throw new WordTruthError(
      `mutool stext extraction failed for ${pdfPath} (exit ${exitCode}): ${stderr.trim()}`,
      "extract",
    );
  }
};

const readCachedGeom = async (geomPath: string, absDocxPath: string): Promise<DocGeom | null> => {
  const file = Bun.file(geomPath);
  if (!(await file.exists())) return null;
  const geom = (await file.json()) as DocGeom;
  return Object.assign({}, geom, {
    file: absDocxPath,
    pages: geom.pages.map((page) =>
      Object.assign({}, page, {
        lines: page.lines.map((line) =>
          Object.assign({}, line, { normText: normalizeLineText(line.text) }),
        ),
      }),
    ),
  });
};

/** Word ground truth for `docxPath`: exports via Word, extracts geometry via
 * mutool, and caches the result under `CACHE_DIR/<sha256>/word-geom.json`.
 * Returns the cached geometry (with `file` rewritten to the requested
 * absolute path) unless `opts.refresh` is set. */
export const getWordTruth = async (
  docxPath: string,
  opts?: { refresh?: boolean },
): Promise<DocGeom> => {
  const absDocxPath = path.resolve(docxPath);
  const sha256 = await sha256OfFile(absDocxPath);
  const dir = cacheDirFor(sha256);
  await mkdir(dir, { recursive: true });

  const geomPath = path.join(dir, GEOM_JSON_FILENAME);
  if (!(opts?.refresh ?? false)) {
    const cached = await readCachedGeom(geomPath, absDocxPath);
    if (cached) return cached;
  }

  if (!(await isWordAvailable())) {
    throw new WordTruthError(
      "Microsoft Word for Mac and/or mutool are not available on this machine.",
      "availability",
    );
  }

  const pdfPath = path.join(dir, PDF_FILENAME);
  await exportViaWord(absDocxPath, pdfPath);

  const xmlPath = path.join(dir, STEXT_XML_FILENAME);
  await runMutoolStext(pdfPath, xmlPath);

  const xml = await Bun.file(xmlPath).text();
  const pages = parseStextXml(xml);

  const [wordVersion, mutoolVersion] = await Promise.all([getWordVersion(), getMutoolVersion()]);
  const geom: DocGeom = {
    source: "word",
    file: absDocxPath,
    pages,
    meta: {
      wordVersion: wordVersion ?? "",
      mutool: mutoolVersion,
      cachedAt: new Date().toISOString(),
      sha256,
    },
  };

  await Bun.write(geomPath, JSON.stringify(geom, null, 2));
  return geom;
};

const listPagePngs = async (pagesDir: string): Promise<string[]> => {
  let entries: string[];
  try {
    entries = await readdir(pagesDir);
  } catch {
    return [];
  }
  return entries
    .filter((name) => PAGE_PNG_RE.test(name))
    .sort((a, b) => pageNumberOfPngName(a) - pageNumberOfPngName(b))
    .map((name) => path.join(pagesDir, name));
};

const pageNumberOfPngName = (filename: string): number => {
  const match = PAGE_PNG_RE.exec(filename);
  return match?.[1] === undefined ? 0 : Number(match[1]);
};

const renderPagePngs = async (
  pdfPath: string,
  pagesDir: string,
  options: { maxPages?: number } = {},
): Promise<void> => {
  const pageRange = options.maxPages === undefined ? [] : [`1-${options.maxPages}`];
  const proc = Bun.spawn(
    [
      "mutool",
      "draw",
      "-r",
      String(PNG_DPI),
      "-o",
      path.join(pagesDir, "p%d.png"),
      pdfPath,
      ...pageRange,
    ],
    { stdout: "pipe", stderr: "pipe" },
  );
  const [exitCode, stderr] = await Promise.all([proc.exited, new Response(proc.stderr).text()]);
  if (exitCode !== 0) {
    throw new WordTruthError(
      `mutool PNG rendering failed for ${pdfPath} (exit ${exitCode}): ${stderr.trim()}`,
      "extract",
    );
  }
};

/** Absolute paths of the cached per-page PNGs for `docxPath`, in page order,
 * rendering them from the cached `word.pdf` (via `getWordTruth`, if needed)
 * at 96dpi on first use. */
export const getWordPagePngs = async (
  docxPath: string,
  options: { maxPages?: number } = {},
): Promise<string[]> => {
  const absDocxPath = path.resolve(docxPath);
  const sha256 = await sha256OfFile(absDocxPath);
  const dir = cacheDirFor(sha256);
  const pdfPath = path.join(dir, PDF_FILENAME);

  if (!(await Bun.file(pdfPath).exists())) {
    await getWordTruth(absDocxPath);
  }

  const pagesDir = path.join(dir, PAGES_DIRNAME);
  await mkdir(pagesDir, { recursive: true });

  const existing = await listPagePngs(pagesDir);
  if (options.maxPages !== undefined && existing.length >= options.maxPages) {
    return existing.slice(0, options.maxPages);
  }
  if (options.maxPages === undefined && existing.length > 0) return existing;

  await renderPagePngs(pdfPath, pagesDir, options);
  const rendered = await listPagePngs(pagesDir);
  return options.maxPages === undefined ? rendered : rendered.slice(0, options.maxPages);
};
