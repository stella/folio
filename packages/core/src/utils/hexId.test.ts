import { describe, expect, spyOn, test } from "bun:test";

import { deterministicHexId, generateHexId, isValidHexId, MAX_HEX_ID_EXCLUSIVE } from "./hexId";

describe("generateHexId", () => {
  test("returns an 8-character uppercase hex string", () => {
    for (let i = 0; i < 50; i++) {
      expect(generateHexId()).toMatch(/^[0-9A-F]{8}$/u);
    }
  });

  test("stays strictly below MAX_HEX_ID_EXCLUSIVE so Word doesn't reject", () => {
    for (let i = 0; i < 200; i++) {
      const id = Number.parseInt(generateHexId(), 16);
      expect(id).toBeGreaterThanOrEqual(1);
      expect(id).toBeLessThan(MAX_HEX_ID_EXCLUSIVE);
    }
  });

  test("never mints 00000000 — Word reads a zero paraId as unassigned", () => {
    const random = spyOn(Math, "random");
    try {
      random.mockReturnValue(0);
      expect(generateHexId()).toBe("00000001");
      // Largest possible Math.random() output still stays below the bound.
      random.mockReturnValue(1 - Number.EPSILON);
      expect(Number.parseInt(generateHexId(), 16)).toBeLessThan(MAX_HEX_ID_EXCLUSIVE);
    } finally {
      random.mockRestore();
    }
  });

  test("produces unique ids in a small sample (sanity check on the rng)", () => {
    const seen = new Set<string>();
    for (let i = 0; i < 100; i++) {
      seen.add(generateHexId());
    }
    expect(seen.size).toBeGreaterThan(95);
  });
});

describe("deterministicHexId", () => {
  test("is stable for the same seed and 8-char uppercase hex", () => {
    const id = deterministicHexId("some:seed");
    expect(id).toMatch(/^[0-9A-F]{8}$/u);
    expect(deterministicHexId("some:seed")).toBe(id);
    expect(deterministicHexId("some:other-seed")).not.toBe(id);
  });

  test("never returns 00000000 and stays below the bound", () => {
    for (let i = 0; i < 500; i++) {
      const id = deterministicHexId(`seed:${i}`);
      expect(id).not.toBe("00000000");
      expect(Number.parseInt(id, 16)).toBeLessThan(MAX_HEX_ID_EXCLUSIVE);
    }
  });
});

describe("isValidHexId", () => {
  test("accepts exactly 8 hex digits", () => {
    expect(isValidHexId("00ABCDEF")).toBe(true);
    expect(isValidHexId(generateHexId())).toBe(true);
  });

  test("rejects malformed ids instead of only escaping them — comment/paragraph paraId must not carry markup", () => {
    expect(isValidHexId('12345678" ><script>alert(1)</script>')).toBe(false);
    expect(isValidHexId("1234567")).toBe(false); // too short
    expect(isValidHexId("123456789")).toBe(false); // too long
    expect(isValidHexId("")).toBe(false);
    expect(isValidHexId(undefined)).toBe(false);
    expect(isValidHexId(null)).toBe(false);
  });
});
