import { createHash } from "node:crypto";
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

/**
 * Playground dev-server port, derived deterministically from the worktree path
 * so parallel git worktrees never collide on one shared port. The old fixed
 * 4200 meant a parity run in worktree A could silently reuse (or fail to
 * displace, under Vite `strictPort`) a stale server started by worktree B —
 * making folio geometry reflect the wrong source. A per-worktree port keeps
 * each worktree's feedback loop isolated and current.
 */
const PORT_BASE = 4200;
const PORT_SPAN = 400;
export const PLAYGROUND_PORT =
  PORT_BASE +
  (Number.parseInt(createHash("sha256").update(REPO_ROOT).digest("hex").slice(0, 8), 16) % PORT_SPAN);
export const PLAYGROUND_URL = `http://localhost:${PLAYGROUND_PORT}`;
/** The dev server reads its port from `FOLIO_PLAYGROUND_PORT` (see the
 * playground `vite.config.ts`); folioExtract sets it when spawning. */
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
