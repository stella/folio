/**
 * Base64 is a total function on byte arrays, and the encoder the package ships
 * has to be total on the runtimes it ships to.
 *
 * The defect this pins is not an edge case in the arithmetic: it is that an
 * encoder can be correct on the runtime its tests run on and throw on the one
 * its users run on. So the properties below drive the portable encoder
 * directly rather than whichever one this runtime selects, cover all 256 byte
 * values and every length residue mod 3, and check the round trip against
 * `atob`, which is the decoder a browser would hand the output to.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { propertyTestTimeout } from "../../../../test/property-testing";

import { registerImage } from "../markdown/images";
import type { RenderContext } from "../markdown/types";
import { bytesToBase64, bytesToDataUrl, encodeBase64Portable } from "./base64";

setDefaultTimeout(propertyTestTimeout(30_000));

const decodeToBytes = (encoded: string): Uint8Array => {
  const binary = atob(encoded);
  const out = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    out[index] = binary.charCodeAt(index);
  }
  return out;
};

const byteArrays = fc.uint8Array({ minLength: 0, maxLength: 600 });

describe("bytes to base64", () => {
  test("every byte array round-trips through atob", () => {
    fc.assert(
      fc.property(byteArrays, (bytes) => {
        expect([...decodeToBytes(encodeBase64Portable(bytes))]).toEqual([...bytes]);
      }),
      { numRuns: 400 },
    );
  });

  test("the portable encoder agrees with the runtime's own", () => {
    fc.assert(
      fc.property(byteArrays, (bytes) => {
        expect(bytesToBase64(bytes)).toBe(encodeBase64Portable(bytes));
      }),
      { numRuns: 400 },
    );
  });

  test("all 256 byte values encode, at every length residue", () => {
    const every = Uint8Array.from({ length: 256 }, (_unused, value) => value);
    for (let length = 0; length <= every.length; length += 1) {
      const slice = every.subarray(0, length);
      const encoded = encodeBase64Portable(slice);
      expect(encoded.length).toBe(Math.ceil(length / 3) * 4);
      expect([...decodeToBytes(encoded)]).toEqual([...slice]);
    }
  });

  test("a data URL carries its media type and payload", () => {
    const bytes = Uint8Array.from([0x89, 0x50, 0x4e, 0x47]);
    expect(bytesToDataUrl(bytes, "image/png")).toBe(
      `data:image/png;base64,${bytesToBase64(bytes)}`,
    );
  });
});

/**
 * `btoa` is only reachable on the path a browser takes, which is the path a
 * runtime carrying `Buffer` hides. Removing the global is what makes the
 * markdown renderer take it here.
 */
const withoutBuffer = <T>(body: () => T): T => {
  const saved = globalThis.Buffer;
  Reflect.deleteProperty(globalThis, "Buffer");
  try {
    return body();
  } finally {
    globalThis.Buffer = saved;
  }
};

describe("markdown image registration", () => {
  test("encodes bytes 0x80-0x9F without a Buffer to fall back on", () => {
    const bytes = Uint8Array.from({ length: 256 }, (_unused, value) => value);
    const context = {
      images: new Map(),
      imagesByPath: new Map(),
      imageCounter: 0,
      opts: {},
    } as unknown as RenderContext;

    const ref = withoutBuffer(() =>
      registerImage(
        context,
        {
          path: "word/media/image1.png",
          filename: "image1.png",
          mimeType: "image/png",
          data: bytes.buffer,
        },
        undefined,
        undefined,
      ),
    );

    expect([...decodeToBytes(ref.base64)]).toEqual([...bytes]);
    expect(ref.dataUrl).toBe(`data:image/png;base64,${ref.base64}`);
  });
});
