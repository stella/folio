import { describe, expect, test } from "bun:test";

import { toArrayBuffer } from "./docxInput";

describe("toArrayBuffer", () => {
  test("reuses an exact Uint8Array backing buffer", async () => {
    const buffer = new ArrayBuffer(4);
    const input = new Uint8Array(buffer);

    const result = await toArrayBuffer(input);

    expect(result).toBe(buffer);
  });

  test("copies only the bytes in a partial Uint8Array view", async () => {
    const buffer = Uint8Array.from([1, 2, 3, 4]).buffer;
    const input = new Uint8Array(buffer, 1, 2);

    const result = await toArrayBuffer(input);

    expect(result).not.toBe(buffer);
    expect([...new Uint8Array(result)]).toEqual([2, 3]);
  });

  test("copies a Uint8Array backed by shared memory", async () => {
    const buffer = new SharedArrayBuffer(3);
    const input = new Uint8Array(buffer);
    input.set([1, 2, 3]);

    const result = await toArrayBuffer(input);

    expect(result).toBeInstanceOf(ArrayBuffer);
    expect([...new Uint8Array(result)]).toEqual([1, 2, 3]);
  });
});
