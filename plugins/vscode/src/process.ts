/**
 * Run the bundled folio CLI in a child process and collect what it printed.
 * Saving goes through here: the CLI does the work outside the extension host,
 * so a large document never stalls every other extension.
 */

import { spawn } from "node:child_process";

import { cliCommand, type CliRuntime } from "./runtime";

export type ProcessResult = {
  readonly code: number | null;
  readonly stdout: string;
  readonly stderr: string;
  readonly timedOut: boolean;
};

export type RunOptions = {
  readonly runtime: CliRuntime;
  readonly args: readonly string[];
  readonly timeoutMs: number;
};

export const runCli = ({ runtime, args, timeoutMs }: RunOptions): Promise<ProcessResult> =>
  new Promise((resolve, reject) => {
    const { command, args: argv, env } = cliCommand(runtime, args);
    const child = spawn(command, [...argv], {
      env: { ...process.env, ...env },
      stdio: ["ignore", "pipe", "pipe"],
      windowsHide: true,
    });
    let stdout = "";
    let stderr = "";
    let timedOut = false;
    child.stdout.setEncoding("utf8").on("data", (chunk: string) => {
      stdout += chunk;
    });
    child.stderr.setEncoding("utf8").on("data", (chunk: string) => {
      stderr += chunk;
    });
    const timer = setTimeout(() => {
      timedOut = true;
      child.kill();
    }, timeoutMs);
    child.on("error", (error) => {
      clearTimeout(timer);
      reject(error);
    });
    child.on("close", (code) => {
      clearTimeout(timer);
      resolve({ code, stdout, stderr, timedOut });
    });
  });

/** The last lines of a crashed process's stderr, for an error message. */
export const stderrTail = (stderr: string, lines = 5): string =>
  stderr.trim().split("\n").slice(-lines).join("\n");
