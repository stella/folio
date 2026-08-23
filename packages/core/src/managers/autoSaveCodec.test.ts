import { describe, expect, test } from "bun:test";
import path from "node:path";

import type { Document } from "../types/document";
import {
  AUTOSAVE_FORMAT_VERSION,
  decodeAutoSaveEnvelope,
  encodeAutoSaveEnvelope,
} from "./autoSaveCodec";

const document = (): Document => ({
  package: {
    document: { content: [] },
    headers: new Map([["rId1", { type: "header", hdrFtrType: "default", content: [] }]]),
    media: new Map([
      [
        "rId2",
        {
          path: "word/media/image1.png",
          mimeType: "image/png",
          data: new Uint8Array([1, 2, 3]).buffer,
        },
      ],
    ]),
    properties: { created: new Date("2026-08-23T10:00:00.000Z") },
  },
});

describe("auto-save codec", () => {
  test("round-trips Maps and Dates without relying on JSON's lossy defaults", () => {
    const encoded = encodeAutoSaveEnvelope(document(), "2026-08-23T11:00:00.000Z");
    const decoded = decodeAutoSaveEnvelope(encoded);

    expect(decoded?.version).toBe(AUTOSAVE_FORMAT_VERSION);
    expect(decoded?.savedAt).toBe("2026-08-23T11:00:00.000Z");
    expect(decoded?.document.package.headers).toBeInstanceOf(Map);
    expect(decoded?.document.package.headers?.get("rId1")?.type).toBe("header");
    expect(decoded?.document.package.properties?.created).toBeInstanceOf(Date);
    expect(
      new Uint8Array(decoded?.document.package.media?.get("rId2")?.data ?? new ArrayBuffer(0)),
    ).toEqual(new Uint8Array([1, 2, 3]));
    expect(encoded).toContain('"type":"arrayBuffer","value":"AQID"');
    expect(encoded).not.toContain('"value":[1,2,3]');
  });

  test("rejects encoded object keys that could mutate a decoded prototype", () => {
    const poisoned =
      '{"document":{"type":"object","value":{"__proto__":{"type":"object","value":{}}}},' +
      `"savedAt":"2026-08-23T11:00:00.000Z","version":${AUTOSAVE_FORMAT_VERSION}}`;

    expect(decodeAutoSaveEnvelope(poisoned)).toBeNull();
  });

  test("keeps binary media near base64 size instead of decimal-array expansion", () => {
    const mediaBytes = 1024 * 1024;
    const largeDocument = document();
    largeDocument.package.media?.set("rId2", {
      path: "word/media/image1.png",
      mimeType: "image/png",
      data: new Uint8Array(mediaBytes).fill(255).buffer,
    });

    const encoded = encodeAutoSaveEnvelope(largeDocument, "2026-08-23T11:00:00.000Z");

    expect(encoded.length).toBeLessThan(mediaBytes * 1.5);
    expect(decodeAutoSaveEnvelope(encoded)).not.toBeNull();
  });

  test("rejects legacy and structurally invalid records before recovery", () => {
    const legacy = JSON.stringify({
      document: { package: { document: { content: [] } } },
      savedAt: "2026-08-23T11:00:00.000Z",
      version: 1,
    });
    const malformed = JSON.stringify({
      document: { type: "object", value: { package: { type: "object", value: {} } } },
      savedAt: "2026-08-23T11:00:00.000Z",
      version: AUTOSAVE_FORMAT_VERSION,
    });

    expect(decodeAutoSaveEnvelope(legacy)).toBeNull();
    expect(decodeAutoSaveEnvelope(malformed)).toBeNull();
  });

  test("keeps the manager from reintroducing direct document JSON serialization", async () => {
    const source = await Bun.file(path.join(import.meta.dir, "AutoSaveManager.ts")).text();
    expect(source).not.toContain("JSON.stringify");
  });
});
