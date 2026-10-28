import { describe, expect, test } from "bun:test";
import {
  closeSync,
  mkdtempSync,
  openSync,
  readSync,
  rmSync,
  writeSync,
} from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

import { systemZipCheck } from "./__tests__/support";
import { ZipArchive, ZipError } from "./archive";
import { createZipStream, zipStream, type ZipStreamEntry } from "./stream";

const collect = async (
  stream: ReadableStream<Uint8Array>,
): Promise<Uint8Array> =>
  new Uint8Array(await new Response(stream).arrayBuffer());

async function* chunked(
  text: string,
  size: number,
): AsyncGenerator<Uint8Array> {
  const bytes = new TextEncoder().encode(text);
  for (let offset = 0; offset < bytes.byteLength; offset += size) {
    yield bytes.subarray(offset, offset + size);
  }
}

describe("streaming writer", () => {
  test("entries stream through and read back", async () => {
    const long = "stream ".repeat(10_000);
    const entries: ZipStreamEntry[] = [
      { path: "folder/" },
      { path: "folder/a.txt", data: chunked(long, 997) },
      { path: "b.bin", data: Uint8Array.of(0, 1, 2, 255) },
      { path: "c/č.txt", data: "text" },
      { path: "empty.txt", data: [] },
    ];
    const bytes = await collect(createZipStream(entries));
    const archive = await ZipArchive.read(bytes);
    expect(archive.paths()).toEqual([
      "folder/a.txt",
      "b.bin",
      "c/č.txt",
      "empty.txt",
    ]);
    expect(await archive.entry("folder/a.txt")?.text()).toBe(long);
    expect(await archive.entry("b.bin")?.bytes()).toEqual(
      Uint8Array.of(0, 1, 2, 255),
    );
    expect(await archive.entry("c/č.txt")?.text()).toBe("text");
    expect(await archive.entry("empty.txt")?.bytes()).toEqual(new Uint8Array());
    if (systemZipCheck.available) {
      expect(systemZipCheck.check(bytes)).toEqual({ ok: true });
    }
  });

  test("entries are pulled only as the output is consumed", async () => {
    const opened: string[] = [];
    async function* lazyEntries(): AsyncGenerator<ZipStreamEntry> {
      for (const name of ["a", "b", "c"]) {
        opened.push(name);
        yield { path: name, data: name };
      }
    }
    const reader = createZipStream(lazyEntries()).getReader();
    await reader.read();
    expect(opened).toEqual(["a"]);
    await reader.cancel();
    expect(opened).toEqual(["a"]);
  });

  test("unsafe and duplicate names fail the stream", async () => {
    for (const entries of [
      [{ path: "../x", data: "x" }],
      [
        { path: "a", data: "x" },
        { path: "a", data: "y" },
      ],
      [{ path: "dir/", data: "x" }],
    ]) {
      // oxlint-disable-next-line no-await-in-loop -- each case is checked in turn
      await expect(collect(createZipStream(entries))).rejects.toBeInstanceOf(
        ZipError,
      );
    }
  });

  test("more than 65535 entries use the Zip64 end records", async () => {
    const count = 65_540;
    const entries = Array.from({ length: count }, (_, index) => ({
      path: `e${index}`,
      data: index % 1000 === 0 ? `${index}` : "",
    }));
    const bytes = await collect(createZipStream(entries));
    const archive = await ZipArchive.read(bytes, {
      limits: { maxEntries: count },
    });
    expect(archive.paths().length).toBe(count);
    expect(await archive.entry("e65000")?.text()).toBe("65000");
    if (systemZipCheck.available) {
      expect(systemZipCheck.check(bytes)).toEqual({ ok: true });
    }
  }, 60_000);

  test("an entry past 4 GiB moves sizes and offsets into Zip64 fields", async () => {
    // Stream 4 GiB + 1 MiB of zeros from one reused chunk into a sparse
    // file: the zero chunks are skipped with a seek, so the disk holds only
    // the records. Info-ZIP and Python then read the whole archive back.
    const zero = new Uint8Array(4 * 1024 * 1024);
    const bigSize = 4 * 1024 * 1024 * 1024 + zero.byteLength;
    async function* zeros(): AsyncGenerator<Uint8Array> {
      for (let written = 0; written < bigSize; written += zero.byteLength)
        yield zero;
    }
    const directory = mkdtempSync(path.join(tmpdir(), "zip64-"));
    const file = path.join(directory, "large.zip");
    const descriptor = openSync(file, "w");
    let offset = 0;
    try {
      for await (const chunk of zipStream([
        { path: "large.bin", data: zeros() },
        { path: "after.txt", data: "after" },
      ])) {
        if (chunk !== zero)
          writeSync(descriptor, chunk, 0, chunk.byteLength, offset);
        offset += chunk.byteLength;
      }
    } finally {
      closeSync(descriptor);
    }
    try {
      expect(offset).toBeGreaterThan(bigSize);
      // The records after the data: central directory, Zip64 end record and
      // locator, end record. Read them back from the file's tail.
      const tailSize = 1024;
      const tail = new Uint8Array(tailSize);
      const reader = openSync(file, "r");
      try {
        readSync(reader, tail, 0, tailSize, offset - tailSize);
      } finally {
        closeSync(reader);
      }
      const view = new DataView(tail.buffer);
      const end = tailSize - 22;
      expect(view.getUint32(end, true)).toBe(0x06_05_4b_50);
      expect(view.getUint32(end + 16, true)).toBe(0xff_ff_ff_ff);
      const locator = end - 20;
      expect(view.getUint32(locator, true)).toBe(0x07_06_4b_50);
      const record =
        Number(view.getBigUint64(locator + 8, true)) - (offset - tailSize);
      expect(view.getUint32(record, true)).toBe(0x06_06_4b_50);
      expect(view.getBigUint64(record + 32, true)).toBe(2n);
      const directoryOffset = Number(view.getBigUint64(record + 48, true));
      expect(directoryOffset).toBeGreaterThan(bigSize);
      // First record: large.bin, both sizes in its Zip64 extra field.
      let cursor = directoryOffset - (offset - tailSize);
      expect(view.getUint32(cursor + 20, true)).toBe(0xff_ff_ff_ff);
      expect(view.getUint32(cursor + 24, true)).toBe(0xff_ff_ff_ff);
      const nameLength = view.getUint16(cursor + 28, true);
      const extra = cursor + 46 + nameLength;
      expect(view.getUint16(extra, true)).toBe(1);
      expect(view.getBigUint64(extra + 4, true)).toBe(BigInt(bigSize));
      expect(view.getBigUint64(extra + 12, true)).toBe(BigInt(bigSize));
      // Second record: after.txt, its local header offset past 4 GiB.
      cursor = extra + 4 + view.getUint16(extra + 2, true);
      expect(view.getUint32(cursor + 42, true)).toBe(0xff_ff_ff_ff);
      const secondExtra = cursor + 46 + view.getUint16(cursor + 28, true);
      expect(view.getBigUint64(secondExtra + 4, true)).toBeGreaterThan(
        BigInt(bigSize),
      );
      if (systemZipCheck.available) {
        // Info-ZIP reads every byte of both entries and checks their CRC-32.
        const unzip = Bun.spawnSync(["unzip", "-tqq", file]);
        expect(unzip.exitCode).toBe(0);
      }
    } finally {
      rmSync(directory, { recursive: true, force: true });
    }
  }, 300_000);
});
