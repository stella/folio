/**
 * Microsoft Word reference renderer: export a DOCX via the locally installed app
 * for Mac (scripted headlessly through AppleScript), extract per-line
 * geometry with `mutool draw -F stext`, and render per-page PNGs for the
 * visual report. Everything is cached under `CACHE_DIR/<sha256 of the docx
 * content>/` so a second call never reopens Word or re-runs mutool.
 */

import { randomUUID } from "node:crypto";
import { mkdir, rm } from "node:fs/promises";
import path from "node:path";

import JSZip from "jszip";

import {
  cacheDirFor,
  extractPdfGeometry,
  getMutoolVersion,
  getPdfPagePngs,
  readCachedGeom,
  sha256OfFile,
} from "./pdfReference";
import type { DocGeom, ReviewView } from "./types";

const WORD_APP_PATH = "/Applications/Microsoft Word.app";
// Large documents can take well over a minute before the app exposes them
// after `open`, so the staged document is polled for up to
// STAGED_DOCUMENT_OPEN_TIMEOUT_SECONDS; the export timeout leaves room for
// that wait plus the PDF save itself.
const STAGED_DOCUMENT_OPEN_TIMEOUT_SECONDS = 120;
const STAGED_DOCUMENT_POLL_SECONDS = 0.5;
const STAGED_DOCUMENT_POLL_ATTEMPTS = Math.ceil(
  STAGED_DOCUMENT_OPEN_TIMEOUT_SECONDS / STAGED_DOCUMENT_POLL_SECONDS,
);
const EXPORT_TIMEOUT_MS = 300_000;
const CLOSE_TIMEOUT_MS = 30_000;
const EXPORT_ATTEMPTS = 2;
const CLOSE_ATTEMPTS = 2;
type WordReviewView = Exclude<ReviewView, "default">;

const DEFAULT_WORD_REVIEW_VIEW: WordReviewView = "final";

const wordArtifactNames = (reviewView: WordReviewView) => ({
  pdf: `word-${reviewView}.pdf`,
  stext: `word-${reviewView}-stext.xml`,
  geom: `word-${reviewView}-geom.json`,
  pages: `word-${reviewView}-pages`,
});

const requireWordReviewView = (reviewView: ReviewView | undefined): WordReviewView => {
  const resolved = reviewView ?? DEFAULT_WORD_REVIEW_VIEW;
  if (resolved === "default") {
    throw new WordTruthError("Microsoft Word requires final or all-markup review view", "export");
  }
  return resolved;
};

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

type BuildExportScriptOptions = {
  docxPath: string;
  pdfPath: string;
  reviewView?: WordReviewView;
};

// Word's scripting dictionary declares a document result for `open`, but
// some current builds return no AppleScript value. Resolve the opened
// document from its unique staged path instead of relying on window focus.
export const buildExportScript = ({
  docxPath,
  pdfPath,
  reviewView = DEFAULT_WORD_REVIEW_VIEW,
}: BuildExportScriptOptions): string => {
  const inFile = escapeAppleScriptString(docxPath);
  const outFile = escapeAppleScriptString(pdfPath);
  const reviewViewScript =
    reviewView === "final"
      ? `set print revisions of theDoc to false
		set revisions view of documentView to revisions view final
		set show revisions and comments of documentView to false`
      : `set print revisions of theDoc to true
		set revisions view of documentView to revisions view final
		set revisions mode of documentView to in line revisions
		set show revisions and comments of documentView to true
		set show insertions and deletions of documentView to true
		set show comments of documentView to false`;
  return `with timeout of ${APPLE_EVENT_TIMEOUT_SECONDS} seconds
	set stagedDocumentPath to "${inFile}"
	set inFile to POSIX file stagedDocumentPath
	tell application "Microsoft Word"
		open inFile
		set theDoc to missing value
		repeat ${STAGED_DOCUMENT_POLL_ATTEMPTS} times
			set openDocuments to {}
			try
				set openDocuments to get every document
			end try
			repeat with candidateDocument in openDocuments
				set candidatePath to missing value
				try
					set candidatePath to POSIX path of ((full name of candidateDocument as text) as alias)
				end try
				if candidatePath is stagedDocumentPath then
					set theDoc to contents of candidateDocument
					exit repeat
				end if
			end repeat
			if theDoc is not missing value then exit repeat
			delay ${STAGED_DOCUMENT_POLL_SECONDS}
		end repeat
		if theDoc is missing value then
			set openDocumentCount to 0
			try
				set openDocumentCount to count of documents
			end try
			error "Word did not expose the staged document within ${STAGED_DOCUMENT_OPEN_TIMEOUT_SECONDS} seconds of opening it (" & openDocumentCount & " documents open)"
		end if
		set documentView to view of active window of theDoc
		${reviewViewScript}
		save as theDoc file name "${outFile}" file format format PDF
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

export const isParityStagedDocumentPath = (filePath: string): boolean =>
  path.dirname(filePath) === WORD_CONTAINER_TMP &&
  /^parity-[a-zA-Z0-9-]+\.docx$/.test(path.basename(filePath));

export const buildCloseStagedDocumentScript = (docxPath: string): string => {
  const stagedDocumentPath = escapeAppleScriptString(docxPath);
  return `with timeout of ${Math.floor(CLOSE_TIMEOUT_MS / 1000)} seconds
	set stagedDocumentPath to "${stagedDocumentPath}"
	tell application "Microsoft Word"
		set openDocuments to get every document
		repeat with candidateDocument in openDocuments
			set candidatePath to missing value
			try
				set candidatePath to POSIX path of ((full name of candidateDocument as text) as alias)
			end try
			if candidatePath is stagedDocumentPath then
				close candidateDocument saving no
				exit repeat
			end if
		end repeat
	end tell
end timeout`;
};

const SETTINGS_PART = "word/settings.xml";
// `w:documentProtection` (ECMA-376 Part 1, 17.15.1.29) restricts editing and
// makes the app reject review-view changes on the document ("Can't set print
// revisions"); `w:writeProtection` (17.15.1.93) can prompt for a password on
// open. Neither changes how content lays out, so the staged copy drops both.
const EDIT_RESTRICTION_ELEMENT_PATTERN =
  /<(?:[A-Za-z_][\w.-]*:)?(documentProtection|writeProtection)\b[^>]*?(?:\/>|>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?\1\s*>)/g;

/** Remove edit restrictions from a settings part, or return undefined when it
 * carries none. */
export const stripEditRestrictions = (settingsXml: string): string | undefined => {
  const stripped = settingsXml.replace(EDIT_RESTRICTION_ELEMENT_PATTERN, "");
  return stripped === settingsXml ? undefined : stripped;
};

/** Bytes to stage for export. Unrestricted documents are staged verbatim; a
 * document with edit restrictions is staged as a copy whose settings part
 * omits them, so both review views can be selected. The source file and its
 * content hash (the cache key) are never changed. */
export const stageableDocxBytes = async (source: Uint8Array): Promise<Uint8Array> => {
  const zip = await JSZip.loadAsync(source);
  const settings = zip.file(SETTINGS_PART);
  if (settings === null) return source;
  const stripped = stripEditRestrictions(await settings.async("string"));
  if (stripped === undefined) return source;
  zip.file(SETTINGS_PART, stripped);
  return await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
};

/** Export `docxPath` to `destPdfPath` by scripting Word. Both sides of the
 * conversion are staged inside Word's sandbox container (see above), and the
 * PDF is moved to `destPdfPath` only on success, so a killed/failed export
 * never leaves a half-written `word.pdf` that a later call would mistake for
 * a valid cache entry. */
const exportViaWord = async (
  docxPath: string,
  destPdfPath: string,
  reviewView: WordReviewView,
): Promise<void> => {
  const stagingToken = `parity-${process.pid}-${Date.now()}-${randomUUID()}`;
  const stagedDocxPath = path.join(WORD_CONTAINER_TMP, `${stagingToken}.docx`);
  const tmpPdfPath = path.join(WORD_CONTAINER_TMP, `${stagingToken}.pdf`);
  await mkdir(WORD_CONTAINER_TMP, { recursive: true });
  await Bun.write(stagedDocxPath, await stageableDocxBytes(await Bun.file(docxPath).bytes()));
  try {
    await runWordExportScript({ docxPath, stagedDocxPath, tmpPdfPath, destPdfPath, reviewView });
  } finally {
    await Promise.all([rm(stagedDocxPath, { force: true }), rm(tmpPdfPath, { force: true })]);
  }
};

type RunWordExportArgs = {
  /** Original path, used only for error messages. */
  docxPath: string;
  stagedDocxPath: string;
  tmpPdfPath: string;
  destPdfPath: string;
  reviewView: WordReviewView;
};

type AppleScriptResult =
  | { status: "success" }
  | { status: "timed-out" }
  | { status: "failed"; exitCode: number; stderr: string };

const runAppleScript = async (script: string, timeoutMs: number): Promise<AppleScriptResult> => {
  const proc = Bun.spawn(["osascript", "-e", script], { stdout: "ignore", stderr: "pipe" });
  let timedOut = false;
  const timer = setTimeout(() => {
    timedOut = true;
    proc.kill();
  }, timeoutMs);

  const [exitCode, stderr] = await Promise.all([proc.exited, new Response(proc.stderr).text()]);
  clearTimeout(timer);

  if (timedOut) return { status: "timed-out" };
  if (exitCode !== 0) return { status: "failed", exitCode, stderr: stderr.trim() };
  return { status: "success" };
};

const describeAppleScriptFailure = (result: Exclude<AppleScriptResult, { status: "success" }>) => {
  if (result.status === "timed-out") return "timed out";
  return `failed (exit ${result.exitCode}): ${result.stderr}`;
};

const closeStagedDocument = async (stagedDocxPath: string): Promise<AppleScriptResult> => {
  if (!isParityStagedDocumentPath(stagedDocxPath)) {
    throw new WordTruthError("Refusing to close a document outside parity staging", "export");
  }
  const script = buildCloseStagedDocumentScript(stagedDocxPath);
  let result = await runAppleScript(script, CLOSE_TIMEOUT_MS);
  for (let attempt = 1; attempt < CLOSE_ATTEMPTS && result.status !== "success"; attempt += 1) {
    result = await runAppleScript(script, CLOSE_TIMEOUT_MS);
  }
  return result;
};

const runWordExportScript = async (args: RunWordExportArgs): Promise<void> => {
  const { docxPath, stagedDocxPath, tmpPdfPath, destPdfPath, reviewView } = args;
  const script = buildExportScript({
    docxPath: stagedDocxPath,
    pdfPath: tmpPdfPath,
    reviewView,
  });

  for (let attempt = 1; attempt <= EXPORT_ATTEMPTS; attempt += 1) {
    const exportResult = await runAppleScript(script, EXPORT_TIMEOUT_MS);
    const closeResult = await closeStagedDocument(stagedDocxPath);

    // A retry is safe only after cleanup confirms that the exact staged
    // document is closed; otherwise a second open can overlap stale state.
    if (closeResult.status !== "success") {
      await rm(tmpPdfPath, { force: true });
      if (exportResult.status !== "success") {
        throw new WordTruthError(
          `Word export ${describeAppleScriptFailure(exportResult)}; cleanup also ${describeAppleScriptFailure(closeResult)} for ${docxPath}`,
          "export",
        );
      }
      throw new WordTruthError(
        `Word cleanup ${describeAppleScriptFailure(closeResult)} for ${docxPath}`,
        "export",
      );
    }

    if (exportResult.status === "success") {
      const exportedFile = Bun.file(tmpPdfPath);
      if ((await exportedFile.exists()) && exportedFile.size > 0) {
        // Cross-directory move: the container tmp dir and the cache dir may sit on
        // different volumes, so copy + remove instead of rename.
        await Bun.write(destPdfPath, exportedFile);
        await rm(tmpPdfPath, { force: true });
        return;
      }
    }

    await rm(tmpPdfPath, { force: true });
    if (attempt < EXPORT_ATTEMPTS) continue;

    if (exportResult.status === "success") {
      throw new WordTruthError(`Word export produced an empty PDF for ${docxPath}`, "export");
    }
    throw new WordTruthError(
      `Word export ${describeAppleScriptFailure(exportResult)} for ${docxPath}`,
      "export",
    );
  }
};

/** Word reference geometry for `docxPath`: exports via Word, extracts via
 * mutool, and caches the result under a view-specific path in
 * `CACHE_DIR/<sha256>`.
 * Returns the cached geometry (with `file` rewritten to the requested
 * absolute path) unless `opts.refresh` is set. */
export const getWordTruth = async (
  docxPath: string,
  opts?: { refresh?: boolean; reviewView?: ReviewView },
): Promise<DocGeom> => {
  const absDocxPath = path.resolve(docxPath);
  const reviewView = requireWordReviewView(opts?.reviewView);
  const artifactNames = wordArtifactNames(reviewView);
  const sha256 = await sha256OfFile(absDocxPath);
  const dir = cacheDirFor(sha256);
  await mkdir(dir, { recursive: true });

  const geomPath = path.join(dir, artifactNames.geom);
  if (!(opts?.refresh ?? false)) {
    const cached = await readCachedGeom({ geomPath, absDocxPath });
    if (cached) return cached;
  }

  if (!(await isWordAvailable())) {
    throw new WordTruthError(
      "Microsoft Word for Mac and/or mutool are not available on this machine.",
      "availability",
    );
  }

  const pdfPath = path.join(dir, artifactNames.pdf);
  await exportViaWord(absDocxPath, pdfPath, reviewView);

  const xmlPath = path.join(dir, artifactNames.stext);
  let pages;
  try {
    pages = await extractPdfGeometry({ pdfPath, xmlPath });
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    throw new WordTruthError(message, "extract");
  }

  await rm(path.join(dir, artifactNames.pages), { recursive: true, force: true });

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
      reviewView,
    },
  };

  await Bun.write(geomPath, JSON.stringify(geom, null, 2));
  return geom;
};

/** Absolute paths of the cached per-page PNGs for `docxPath`, in page order,
 * rendering them from the matching view's cached PDF (via `getWordTruth`, if
 * needed) at 96dpi on first use. */
export const getWordPagePngs = async (
  docxPath: string,
  options: { maxPages?: number; reviewView?: ReviewView } = {},
): Promise<string[]> => {
  const absDocxPath = path.resolve(docxPath);
  const reviewView = requireWordReviewView(options.reviewView);
  const artifactNames = wordArtifactNames(reviewView);
  const sha256 = await sha256OfFile(absDocxPath);
  const dir = cacheDirFor(sha256);
  const pdfPath = path.join(dir, artifactNames.pdf);

  if (!(await Bun.file(pdfPath).exists())) {
    await getWordTruth(absDocxPath, { refresh: true, reviewView });
  }

  const pagesDir = path.join(dir, artifactNames.pages);
  await mkdir(pagesDir, { recursive: true });
  try {
    return await getPdfPagePngs({
      pdfPath,
      pagesDir,
      ...(options.maxPages === undefined ? {} : { maxPages: options.maxPages }),
    });
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    throw new WordTruthError(message, "extract");
  }
};
