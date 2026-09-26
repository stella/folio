/**
 * Run the bundled folio CLI in a child process and collect what it printed.
 * `render` and `save` both go through here.
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
  readonly signal?: AbortSignal;
  readonly timeoutMs: number;
};

export const runCli = ({ runtime, args, signal, timeoutMs }: RunOptions): Promise<ProcessResult> =>
  new Promise((resolve, reject) => {
    const { command, args: argv, env } = cliCommand(runtime, args);
    const child = spawn(command, [...argv], {
      env: { ...process.env, ...env },
      stdio: ["ignore", "pipe", "pipe"],
      ...(signal !== undefined && { signal }),
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
