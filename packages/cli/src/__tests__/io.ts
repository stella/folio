import type { FolioCliIo } from "../cli";

export type CapturedIo = { io: FolioCliIo; stdout: () => string; stderr: () => string };

/**
 * An environment where git has no user configuration, so author resolution
 * sees only what a test passes.
 */
export const ISOLATED_GIT_ENV: Readonly<Record<string, string | undefined>> = {
  PATH: process.env["PATH"],
  GIT_CONFIG_GLOBAL: "/dev/null",
  GIT_CONFIG_NOSYSTEM: "1",
};

type CaptureIoOptions = {
  isTTY?: boolean;
  stdin?: string;
  env?: Readonly<Record<string, string | undefined>>;
  cwd?: string;
};

/** A {@link FolioCliIo} that records what the command prints. */
export const captureIo = ({
  isTTY = false,
  stdin = "",
  env = { ...ISOLATED_GIT_ENV, FOLIO_AUTHOR: "Test Reviewer" },
  cwd = "/",
}: CaptureIoOptions = {}): CapturedIo => {
  let stdout = "";
  let stderr = "";
  return {
    io: {
      stdout: (text) => {
        stdout += text;
      },
      stderr: (text) => {
        stderr += text;
      },
      readStdin: () => Promise.resolve(stdin),
      isTTY,
      env,
      cwd,
    },
    stdout: () => stdout,
    stderr: () => stderr,
  };
};

/** The parsed envelope a JSON-mode command printed. */
export const envelopeOf = (text: string): Record<string, unknown> => {
  const parsed: unknown = JSON.parse(text);
  if (typeof parsed !== "object" || parsed === null) throw new Error("not an envelope");
  return { ...parsed };
};

/** The `data` of a successful envelope, or a thrown error naming the failure. */
export const dataOf = (text: string): Record<string, unknown> => {
  const envelope = envelopeOf(text);
  const data = envelope["data"];
  if (envelope["ok"] !== true || typeof data !== "object" || data === null) {
    throw new Error(`expected success, got ${text}`);
  }
  return { ...data };
};
