import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { streamZipCodec } from "./codec";

setDefaultTimeout(propertyTestTimeout(60_000));

/** The CRC-32 the web codecs check every inflated entry with. */
const webCrc32 = streamZipCodec("gzip").crc32;

test("the web codec's CRC-32 matches the runtime's, whole and chained", () => {
  assertProperty(
    fc.property(fc.uint8Array({ maxLength: 600 }), fc.nat(), (data, at) => {
      const want = Bun.hash.crc32(data) >>> 0;
      expect(webCrc32(data)).toBe(want);
      const cut = data.length === 0 ? 0 : at % (data.length + 1);
      expect(
        webCrc32(data.subarray(cut), webCrc32(data.subarray(0, cut))),
      ).toBe(want);
    }),
    {
      numRuns: 500,
      id: "the web codec's CRC-32 matches the runtime's, whole and chained",
    },
  );
});

/** One byte per step over the classic table: the shape that made opening slow. */
const byteByByteCrc32 = (data: Uint8Array): number => {
  const table = new Uint32Array(256);
  for (let index = 0; index < 256; index++) {
    let value = index;
    for (let bit = 0; bit < 8; bit++)
      value = value & 1 ? 0xed_b8_83_20 ^ (value >>> 1) : value >>> 1;
    table[index] = value >>> 0;
  }
  let crc = 0xff_ff_ff_ff;
  for (const byte of data) crc = table[(crc ^ byte) & 0xff]! ^ (crc >>> 8);
  return ~crc >>> 0;
};

/** The least CPU time of three runs of `run`, in microseconds. */
const cpuMicros = (run: () => void): number => {
  let least = Number.POSITIVE_INFINITY;
  for (let round = 0; round < 3; round++) {
    const start = process.cpuUsage();
    run();
    const spent = process.cpuUsage(start);
    least = Math.min(least, spent.user + spent.system);
  }
  return least;
};

test("the web codec checks an entry's CRC-32 several times faster than byte by byte", () => {
  // Every entry a browser inflates is checked here, so a slow loop is paid on
  // every open: CPU time against a byte-by-byte loop on the same engine and
  // the same bytes, which holds on a loaded machine where wall time does not.
  const data = new Uint8Array(16 << 20);
  for (let offset = 0; offset < data.length; offset += 65_536)
    crypto.getRandomValues(data.subarray(offset, offset + 65_536));
  expect(webCrc32(data)).toBe(byteByByteCrc32(data));
  const fast = cpuMicros(() => webCrc32(data));
  const slow = cpuMicros(() => byteByByteCrc32(data));
  expect(slow / fast).toBeGreaterThan(4);
});
