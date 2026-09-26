import { describe, expect, test } from "bun:test";

import { decodeBackup, encodeBackup } from "./backup";

const BASELINE = "c".repeat(64);

describe("backup", () => {
  test("round-trips the baseline and the bytes, newlines and all", () => {
    const bytes = new Uint8Array([0x50, 0x4b, 0x0a, 0x00, 0x0a, 0xff]);

    const decoded = decodeBackup(encodeBackup({ baseline: BASELINE, bytes }));

    expect(decoded).toEqual({ baseline: BASELINE, bytes });
  });

  test("starts with a readable header", () => {
    const encoded = encodeBackup({ baseline: BASELINE, bytes: new Uint8Array([1]) });

    expect(new TextDecoder().decode(encoded.subarray(0, encoded.indexOf(0x0a)))).toBe(
      `folio-vscode-backup {"version":1,"baseline":"${BASELINE}"}`,
    );
  });

  test("refuses data it did not write", () => {
    const encode = (text: string) => new TextEncoder().encode(text);

    expect(decodeBackup(new Uint8Array([0x50, 0x4b, 0x03, 0x04]))).toBeNull();
    expect(decodeBackup(encode("folio-vscode-backup {\n"))).toBeNull();
    expect(decodeBackup(encode('folio-vscode-backup {"version":2,"baseline":"x"}\n'))).toBeNull();
    expect(
      decodeBackup(encode('folio-vscode-backup {"version":1,"baseline":"not-a-version"}\nPK')),
    ).toBeNull();
    expect(decodeBackup(encode(`folio-vscode-backup ${"x".repeat(5000)}\n`))).toBeNull();
  });
});
