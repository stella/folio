/** Run the packed `folio` executable as a user's shell would. */

import { execFile } from "node:child_process";
import { realpathSync } from "node:fs";
import path from "node:path";

/**
 * The `folio` executable the install linked, as `npx folio` finds it. The
 * scenarios run from the consumer project's root.
 */
const FOLIO_BIN = realpathSync(path.resolve("node_modules", ".bin", "folio"));

export type FolioRun = {
  code: number;
  /** The parsed JSON envelope (`{ ok, data }` or `{ ok: false, error }`). */
  json: { ok: boolean; data?: Record<string, unknown>; error?: { code: string; message: string } };
  stderr: string;
};

export const folio = (args: readonly string[], cwd: string): Promise<FolioRun> =>
  new Promise((resolve, reject) => {
    execFile(
      process.execPath,
      [FOLIO_BIN, ...args, "--output", "json"],
      { cwd, env: { ...process.env, FOLIO_AUTHOR: "CLI Scenario" }, maxBuffer: 64 * 1024 * 1024 },
      (error, stdout, stderr) => {
        const code = error && typeof error.code === "number" ? error.code : 0;
        try {
          resolve({ code, json: JSON.parse(stdout) as FolioRun["json"], stderr });
        } catch {
          reject(
            new Error(`folio ${args.join(" ")} printed no JSON (exit ${code}): ${stdout}${stderr}`),
          );
        }
      },
    );
  });
