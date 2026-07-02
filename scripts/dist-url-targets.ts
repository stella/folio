// Scan built dist JS for `new URL("<rel>", import.meta.url)` references and
// report the ones whose target does not exist on disk. Shared by
// `validate-dist.ts`; kept as a pure helper so the resolution rules (skip
// absolute URLs, strip `?query`/`#hash` before the existence check) are unit
// tested in `dist-url-targets.test.ts`, mirroring the oxlint rule
// (`.oxlint-plugins/folio-asset-urls.ts`) and the architecture test.

import { existsSync } from "node:fs";
import path from "node:path";

export type DistJsFile = { file: string; code: string };

export type UrlTargetScan = {
  /** Relative `new URL(..., import.meta.url)` references found. */
  total: number;
  /** `"<file> -> <spec>"` entries whose resolved target is missing on disk. */
  dangling: string[];
};

const URL_TARGET_RE = /new URL\(\s*["']([^"']+)["']\s*,\s*import\.meta\.url\s*\)/gu;

export const scanDistUrlTargets = (files: DistJsFile[], distDir: string): UrlTargetScan => {
  const dangling: string[] = [];
  let total = 0;
  for (const { file, code } of files) {
    for (const match of code.matchAll(URL_TARGET_RE)) {
      const spec = match[1] ?? "";
      // Only file-relative specifiers resolve against the module; skip absolute
      // URLs (`https:`, `data:`, `blob:`) and protocol-relative ones.
      if (/^[a-z][a-z0-9+.-]*:/iu.test(spec) || spec.startsWith("//")) {
        continue;
      }
      total += 1;
      // A `?query` / `#hash` suffix (e.g. `./x.worker.js?worker`) is bundler
      // routing, not part of the on-disk filename; strip it before resolving.
      const specPath = spec.split(/[?#]/u)[0] ?? spec;
      const resolved = path.resolve(path.dirname(path.join(distDir, file)), specPath);
      if (!existsSync(resolved)) {
        dangling.push(`${file} -> ${spec}`);
      }
    }
  }
  return { total, dangling };
};
