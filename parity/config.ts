import path from "node:path";
import type { ComparisonTolerances } from "./types";

export const PARITY_ROOT = import.meta.dir;
export const REPO_ROOT = path.resolve(PARITY_ROOT, "..");

/** Word ground-truth artifacts, keyed by sha256 of the .docx content. */
export const CACHE_DIR = path.join(PARITY_ROOT, ".cache");
/** Generated HTML report output. */
export const REPORT_DIR = path.join(PARITY_ROOT, "report");

/** Playground fixture dir served by the Vite middleware at /fixtures/<name>. */
export const FIXTURES_DIR = path.join(REPO_ROOT, "tests", "visual", "fixtures");
/** Prefix for docx files staged temporarily into FIXTURES_DIR so the
 * playground can load arbitrary inputs; gitignored and always cleaned up. */
export const TMP_FIXTURE_PREFIX = "parity-tmp-";

export const PLAYGROUND_URL = "http://localhost:4200";
export const PLAYGROUND_DEV_COMMAND = ["bun", "--filter", "@stll/playground", "dev"];

/** Default corpus scanned when the CLI gets no explicit paths. */
export const DEFAULT_CORPUS_DIRS = [
  FIXTURES_DIR,
  path.join(REPO_ROOT, "packages", "core", "src", "docx", "__tests__", "__fixtures__"),
];

export const DEFAULT_TOLERANCES: ComparisonTolerances = {
  xPt: 1.5,
  yResidualPt: 2.5,
  widthPt: 2.0,
  widthRelative: 0.015,
};

export const PX_TO_PT = 72 / 96;
