/**
 * LibreOffice reference renderer: convert DOCX to PDF with an isolated,
 * headless Writer process, then extract the shared comparison geometry.
 */

import { mkdir, mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { pathToFileURL } from "node:url";

import {
  cacheDirFor,
  extractPdfGeometry,
  getMutoolVersion,
  getPdfPagePngs,
  readCachedGeom,
  sha256OfFile,
} from "./pdfReference";
import type { DocGeom } from "./types";

const LIBREOFFICE_APP_BINARY = "/Applications/LibreOffice.app/Contents/MacOS/soffice";
const EXPORT_TIMEOUT_MS = 180_000;

const PDF_FILENAME = "libreoffice.pdf";
const STEXT_XML_FILENAME = "libreoffice-stext.xml";
const GEOM_JSON_FILENAME = "libreoffice-geom.json";
const PAGES_DIRNAME = "libreoffice-pages";

export type LibreOfficeTruthStage = "availability" | "export" | "extract";

export class LibreOfficeTruthError extends Error {
  readonly stage: LibreOfficeTruthStage;

  constructor(message: string, stage: LibreOfficeTruthStage) {
    super(message);
    this.name = "LibreOfficeTruthError";
    this.stage = stage;
  }
}

const getLibreOfficeBinary = async (): Promise<string | null> => {
  const fromPath = Bun.which("soffice") ?? Bun.which("libreoffice");
  if (fromPath !== null) return fromPath;
  return (await Bun.file(LIBREOFFICE_APP_BINARY).exists()) ? LIBREOFFICE_APP_BINARY : null;
};

export const isLibreOfficeAvailable = async (): Promise<boolean> =>
  (await getLibreOfficeBinary()) !== null && Bun.which("mutool") !== null;

let cachedLibreOfficeVersion: string | null | undefined;

export const getLibreOfficeVersion = async (): Promise<string | null> => {
  if (cachedLibreOfficeVersion !== undefined) return cachedLibreOfficeVersion;

  const binary = await getLibreOfficeBinary();
  if (binary === null) {
    cachedLibreOfficeVersion = null;
    return cachedLibreOfficeVersion;
  }

  const proc = Bun.spawn([binary, "--headless", "--version"], {
    stdout: "pipe",
    stderr: "pipe",
  });
  const [exitCode, stdout] = await Promise.all([proc.exited, new Response(proc.stdout).text()]);
  const version = stdout.trim().replace(/^LibreOffice\s+/u, "");
  cachedLibreOfficeVersion = exitCode === 0 && version !== "" ? version : null;
  return cachedLibreOfficeVersion;
};

type BuildLibreOfficeExportArgsOptions = {
  binary: string;
  inputPath: string;
  outputDir: string;
  profileDir: string;
};

export const buildLibreOfficeExportArgs = ({
  binary,
  inputPath,
  outputDir,
  profileDir,
}: BuildLibreOfficeExportArgsOptions): string[] => [
  binary,
  "--headless",
  "--nologo",
  "--nodefault",
  "--nofirststartwizard",
  "--norestore",
  `-env:UserInstallation=${pathToFileURL(profileDir).href}`,
  "--convert-to",
  "pdf:writer_pdf_Export",
  "--outdir",
  outputDir,
  inputPath,
];

const exportViaLibreOffice = async (docxPath: string, destPdfPath: string): Promise<void> => {
  const binary = await getLibreOfficeBinary();
  if (binary === null) {
    throw new LibreOfficeTruthError(
      "LibreOffice is not available on this machine.",
      "availability",
    );
  }

  const workDir = await mkdtemp(path.join(tmpdir(), "folio-libreoffice-"));
  const inputPath = path.join(workDir, "input.docx");
  const outputPath = path.join(workDir, "input.pdf");
  const profileDir = path.join(workDir, "profile");
  await mkdir(profileDir, { recursive: true });
  await Bun.write(inputPath, Bun.file(docxPath));

  try {
    const proc = Bun.spawn(
      buildLibreOfficeExportArgs({
        binary,
        inputPath,
        outputDir: workDir,
        profileDir,
      }),
      // LibreOffice may hand conversion to a child process that keeps inherited
      // pipe descriptors open after the launcher exits. Ignoring output avoids
      // waiting forever for EOF after a successful conversion.
      { stdout: "ignore", stderr: "ignore" },
    );

    let timedOut = false;
    const timeout = setTimeout(() => {
      timedOut = true;
      proc.kill();
    }, EXPORT_TIMEOUT_MS);
    const exitCode = await proc.exited;
    clearTimeout(timeout);

    if (timedOut) {
      throw new LibreOfficeTruthError(
        `LibreOffice export timed out after ${EXPORT_TIMEOUT_MS}ms for ${docxPath}`,
        "export",
      );
    }
    if (exitCode !== 0) {
      throw new LibreOfficeTruthError(
        `LibreOffice export failed for ${docxPath} (exit ${exitCode})`,
        "export",
      );
    }

    const output = Bun.file(outputPath);
    if (!(await output.exists()) || output.size === 0) {
      throw new LibreOfficeTruthError(
        `LibreOffice export produced no PDF for ${docxPath}`,
        "export",
      );
    }
    await Bun.write(destPdfPath, output);
  } finally {
    await rm(workDir, { recursive: true, force: true });
  }
};

export const getLibreOfficeTruth = async (
  docxPath: string,
  options: { refresh?: boolean } = {},
): Promise<DocGeom> => {
  const absDocxPath = path.resolve(docxPath);
  const sha256 = await sha256OfFile(absDocxPath);
  const dir = cacheDirFor(sha256);
  await mkdir(dir, { recursive: true });

  const geomPath = path.join(dir, GEOM_JSON_FILENAME);
  if (!(options.refresh ?? false)) {
    const cached = await readCachedGeom(geomPath, absDocxPath);
    if (cached) return cached;
  }

  if (!(await isLibreOfficeAvailable())) {
    throw new LibreOfficeTruthError(
      "LibreOffice and/or mutool are not available on this machine.",
      "availability",
    );
  }

  const pdfPath = path.join(dir, PDF_FILENAME);
  await exportViaLibreOffice(absDocxPath, pdfPath);

  const xmlPath = path.join(dir, STEXT_XML_FILENAME);
  let pages;
  try {
    pages = await extractPdfGeometry(pdfPath, xmlPath);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    throw new LibreOfficeTruthError(message, "extract");
  }

  const [libreOfficeVersion, mutoolVersion] = await Promise.all([
    getLibreOfficeVersion(),
    getMutoolVersion(),
  ]);
  const geom: DocGeom = {
    source: "libreoffice",
    file: absDocxPath,
    pages,
    meta: {
      libreOfficeVersion: libreOfficeVersion ?? "",
      mutool: mutoolVersion,
      cachedAt: new Date().toISOString(),
      sha256,
    },
  };

  await Bun.write(geomPath, JSON.stringify(geom, null, 2));
  return geom;
};

export const getLibreOfficePagePngs = async (
  docxPath: string,
  options: { maxPages?: number } = {},
): Promise<string[]> => {
  const absDocxPath = path.resolve(docxPath);
  const sha256 = await sha256OfFile(absDocxPath);
  const dir = cacheDirFor(sha256);
  const pdfPath = path.join(dir, PDF_FILENAME);

  if (!(await Bun.file(pdfPath).exists())) {
    await getLibreOfficeTruth(absDocxPath);
  }

  const pagesDir = path.join(dir, PAGES_DIRNAME);
  await mkdir(pagesDir, { recursive: true });
  return await getPdfPagePngs({
    pdfPath,
    pagesDir,
    ...(options.maxPages === undefined ? {} : { maxPages: options.maxPages }),
  });
};
