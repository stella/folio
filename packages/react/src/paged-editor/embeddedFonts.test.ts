import { afterEach, beforeEach, describe, expect, mock, test } from "bun:test";

import type { EmbeddedFont } from "@stll/folio-core/fonts/embeddedFonts";

// Stub the headless extractor so this test targets only the DOM-registration
// guard logic; the real extract/de-obfuscate path is covered in folio-core.
let extractImpl: (buffer: ArrayBuffer) => Promise<EmbeddedFont[]>;
void mock.module("@stll/folio-core/fonts/embeddedFonts", () => ({
  extractEmbeddedFonts: (buffer: ArrayBuffer) => extractImpl(buffer),
}));

const { loadEmbeddedFontFaces } = await import("./embeddedFonts");

function fakeFont(): EmbeddedFont {
  return {
    family: "My Brand Sans",
    style: "normal",
    weight: 400,
    bytes: new Uint8Array([0x00, 0x01, 0x00, 0x00]),
    subsetted: false,
  };
}

type FontFaceMode = "load-ok" | "load-reject" | "ctor-throw";

let fontFaceMode: FontFaceMode;
let addedFaces: unknown[];
let deletedFaces: unknown[];

class FakeFontFace {
  family: string;

  constructor(family: string, _source: unknown, _descriptors: unknown) {
    if (fontFaceMode === "ctor-throw") {
      throw new SyntaxError("invalid font family");
    }
    this.family = family;
  }

  load(): Promise<FakeFontFace> {
    if (fontFaceMode === "load-reject") {
      return Promise.reject(new Error("load failed"));
    }
    return Promise.resolve(this);
  }
}

type StubbedGlobal = "document" | "FontFace";

// CI runners expose `document`/`FontFace` as readonly globals, so a plain
// assignment throws. Redefine them via a configurable descriptor (matching
// folio-core's global-stub tests) and delete/restore afterwards.
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
  } else {
    Reflect.deleteProperty(globalThis, prop);
  }
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
  originalDocument = (globalThis as { document?: unknown }).document;
  hadFontFace = "FontFace" in globalThis;
  originalFontFace = (globalThis as { FontFace?: unknown }).FontFace;

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

describe("loadEmbeddedFontFaces (best-effort)", () => {
  test("registers and returns the loaded faces on success", async () => {
    const faces = await loadEmbeddedFontFaces(new ArrayBuffer(8));
    expect(faces).toHaveLength(1);
    expect(addedFaces).toHaveLength(1);
    expect(deletedFaces).toHaveLength(0);
  });

  test("does not throw and returns [] when extraction throws", async () => {
    extractImpl = () => Promise.reject(new Error("corrupt zip"));
    expect(await loadEmbeddedFontFaces(new ArrayBuffer(8))).toEqual([]);
    expect(addedFaces).toHaveLength(0);
  });

  test("does not throw and skips a face when FontFace construction throws", async () => {
    fontFaceMode = "ctor-throw";
    expect(await loadEmbeddedFontFaces(new ArrayBuffer(8))).toEqual([]);
    expect(addedFaces).toHaveLength(0);
  });

  test("drops a face whose load() rejects", async () => {
    fontFaceMode = "load-reject";
    expect(await loadEmbeddedFontFaces(new ArrayBuffer(8))).toEqual([]);
    expect(addedFaces).toHaveLength(1);
    expect(deletedFaces).toHaveLength(1);
  });

  test("returns [] outside a DOM (no document.fonts)", async () => {
    stubGlobal("document", undefined);
    expect(await loadEmbeddedFontFaces(new ArrayBuffer(8))).toEqual([]);
  });
});
