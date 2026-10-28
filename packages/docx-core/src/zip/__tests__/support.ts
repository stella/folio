import { spawnSync } from "node:child_process";
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

/** One entry as a deliberately naive walk of the central directory sees it. */
export type ZipRecord = {
  path: string;
  flags: number;
  method: number;
  crc32: number;
  dosDateTime: { date: number; time: number };
  /** The entry's stored (compressed) bytes. */
  data: Uint8Array;
};

/**
 * Walk a classic (non-Zip64) archive's central directory. Independent of the
 * module under test, so tests can assert on the bytes a writer produced.
 */
export const readZipRecords = (bytes: Uint8Array): ZipRecord[] => {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const end = bytes.byteLength - 22;
  if (view.getUint32(end, true) !== 0x06_05_4b_50) {
    throw new Error("expected an end record without a comment");
  }
  const count = view.getUint16(end + 10, true);
  let cursor = view.getUint32(end + 16, true);
  const records: ZipRecord[] = [];
  for (let index = 0; index < count; index++) {
    const nameLength = view.getUint16(cursor + 28, true);
    const local = view.getUint32(cursor + 42, true);
    const compressedSize = view.getUint32(cursor + 20, true);
    const dataStart =
      local +
      30 +
      view.getUint16(local + 26, true) +
      view.getUint16(local + 28, true);
    records.push({
      path: new TextDecoder().decode(
        bytes.subarray(cursor + 46, cursor + 46 + nameLength),
      ),
      flags: view.getUint16(cursor + 8, true),
      method: view.getUint16(cursor + 10, true),
      crc32: view.getUint32(cursor + 16, true),
      dosDateTime: {
        time: view.getUint16(cursor + 12, true),
        date: view.getUint16(cursor + 14, true),
      },
      data: bytes.slice(dataStart, dataStart + compressedSize),
    });
    cursor +=
      46 +
      nameLength +
      view.getUint16(cursor + 30, true) +
      view.getUint16(cursor + 32, true);
  }
  return records;
};

const PYTHON_CHECK = [
  "import sys, zipfile",
  "with zipfile.ZipFile(sys.argv[1]) as archive:",
  "    bad = archive.testzip()",
  "    sys.exit(1 if bad else 0)",
].join("\n");

type SystemCheck = { ok: true } | { ok: false; tool: string; output: string };

const run = (command: string, args: string[]): SystemCheck => {
  const result = spawnSync(command, args, { encoding: "utf8" });
  return result.status === 0
    ? { ok: true }
    : { ok: false, tool: command, output: `${result.stdout}${result.stderr}` };
};

/**
 * Independent readers: Info-ZIP `unzip -t` tests every entry's CRC, and
 * Python's `zipfile` re-reads the archive with its own parser.
 */
export const systemZipCheck = {
  available: Bun.which("unzip") !== null && Bun.which("python3") !== null,
  check: (bytes: Uint8Array): SystemCheck => {
    const directory = mkdtempSync(path.join(tmpdir(), "zip-check-"));
    const file = path.join(directory, "archive.zip");
    try {
      writeFileSync(file, bytes);
      const unzip = run("unzip", ["-tqq", file]);
      if (!unzip.ok) return unzip;
      return run("python3", ["-c", PYTHON_CHECK, file]);
    } finally {
      rmSync(directory, { recursive: true, force: true });
    }
  },
};
