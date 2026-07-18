import { afterEach, beforeEach, describe, expect, mock, test } from "bun:test";

import type { EmbeddedFont } from "@stll/folio-core/fonts/embeddedFonts";

// The mock replaces the module registry entry for every later importer in the
// test process, so it must keep the real exports and only stub extraction.
const actualEmbeddedFonts = await import("@stll/folio-core/fonts/embeddedFonts");

let extractImpl: (buffer: ArrayBuffer) => Promise<EmbeddedFont[]>;
void mock.module("@stll/folio-core/fonts/embeddedFonts", () => ({
  ...actualEmbeddedFonts,
  extractEmbeddedFonts: (buffer: ArrayBuffer) => extractImpl(buffer),
}));

const { loadEmbeddedFontFaces } = await import("./embeddedFonts");

function fakeFont(): EmbeddedFont {
  return {
    family: "folio-embedded-test-nonce-Document Sans",
    originalFamily: "Document Sans",
    style: "normal",
    weight: 400,
    bytes: new Uint8Array([0x00, 0x01, 0x00, 0x00]),
    subsetted: false,
  };
}

type FontFaceMode = "load-ok" | "load-reject" | "ctor-throw";
type StubbedGlobal = "document" | "FontFace";

let fontFaceMode: FontFaceMode;
let addedFaces: unknown[];
let deletedFaces: unknown[];

class FakeFontFace {
  constructor(_family: string, _source: unknown, _descriptors: unknown) {
    if (fontFaceMode === "ctor-throw") {
      throw new SyntaxError("invalid font family");
    }
  }

  load(): Promise<FakeFontFace> {
    if (fontFaceMode === "load-reject") {
      return Promise.reject(new Error("load failed"));
    }
    return Promise.resolve(this);
  }
}

function stubGlobal(prop: StubbedGlobal, value: unknown): void {
  Object.defineProperty(globalThis, prop, { value, configurable: true, writable: true });
}

function restoreGlobal(prop: StubbedGlobal, had: boolean, original: unknown): void {
  if (had) {
    Object.defineProperty(globalThis, prop, {
      value: original,
      configurable: true,
      writable: true,
    });
    return;
  }
  Reflect.deleteProperty(globalThis, prop);
}

let hadDocument: boolean;
let originalDocument: unknown;
let hadFontFace: boolean;
let originalFontFace: unknown;

beforeEach(() => {
  fontFaceMode = "load-ok";
  addedFaces = [];
  deletedFaces = [];
  extractImpl = () => Promise.resolve([fakeFont()]);

  hadDocument = "document" in globalThis;
  originalDocument = Reflect.get(globalThis, "document");
  hadFontFace = "FontFace" in globalThis;
  originalFontFace = Reflect.get(globalThis, "FontFace");

  stubGlobal("document", {
    fonts: {
      add: (face: unknown) => addedFaces.push(face),
      delete: (face: unknown) => {
        deletedFaces.push(face);
        return true;
      },
    },
  });
  stubGlobal("FontFace", FakeFontFace);
});

afterEach(() => {
  restoreGlobal("document", hadDocument, originalDocument);
  restoreGlobal("FontFace", hadFontFace, originalFontFace);
});

describe("loadEmbeddedFontFaces", () => {
  test("registers faces extracted from the document", async () => {
    const { faces, familyMap } = await loadEmbeddedFontFaces(new ArrayBuffer(8));
    expect(faces).toHaveLength(1);
    expect(addedFaces).toHaveLength(1);
    expect(deletedFaces).toHaveLength(0);
    expect(familyMap.get("Document Sans")).toBe("folio-embedded-test-nonce-Document Sans");
  });

  test("falls back when extraction fails", async () => {
    extractImpl = () => Promise.reject(new Error("corrupt package"));
    expect(await loadEmbeddedFontFaces(new ArrayBuffer(8))).toEqual({
      faces: [],
      familyMap: new Map(),
    });
    expect(addedFaces).toHaveLength(0);
  });

  test("skips a face rejected by the browser", async () => {
    fontFaceMode = "load-reject";
    const { faces } = await loadEmbeddedFontFaces(new ArrayBuffer(8));
    expect(faces).toEqual([]);
    expect(addedFaces).toHaveLength(1);
    expect(deletedFaces).toHaveLength(1);
  });

  test("skips invalid face descriptors", async () => {
    fontFaceMode = "ctor-throw";
    const { faces } = await loadEmbeddedFontFaces(new ArrayBuffer(8));
    expect(faces).toEqual([]);
    expect(addedFaces).toHaveLength(0);
  });

  test("is a no-op outside a browser document", async () => {
    stubGlobal("document", undefined);
    expect(await loadEmbeddedFontFaces(new ArrayBuffer(8))).toEqual({
      faces: [],
      familyMap: new Map(),
    });
  });
});
