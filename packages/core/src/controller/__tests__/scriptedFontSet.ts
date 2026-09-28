/**
 * A scripted stand-in for the browser's `document.fonts`, with bundled faces
 * split into `unicode-range` subsets the way fontsource ships them, and a
 * canvas whose widths depend on which subsets have loaded.
 *
 * The browser behaviours that matter to layout are modelled, and nothing else:
 * - `load(font, text)` fetches only the faces whose range covers `text` (a
 *   space when omitted), so a call without the document's text loads the
 *   Latin subset alone;
 * - painted text requests the subsets it needs on its own (`requestForText`);
 * - a face's status flips when its load completes, while `loadingdone` is
 *   queued and dispatched only when the test delivers it, so a layout can run
 *   between the two in either order;
 * - a canvas measures each character in the first family of its stack with a
 *   loaded face covering it, else in a fallback of a different width.
 *
 * The test decides the order and interleaving of completions and deliveries.
 */

import type { LayoutFontSet, LayoutFontSetFace } from "../fontReadiness";
import { parseFontFamilyList } from "../../utils/fontResolver";

// Synthetic text, one entry per script a bundled subset covers; the Greek and
// Cyrillic entries reach into the `-ext` subsets too.
export const SCRIPTS = ["english", "czech", "polish", "greek", "cyrillic"] as const;
export type Script = (typeof SCRIPTS)[number];
export const SCRIPT_TEXT = {
  english: "The Supplier shall deliver the goods within ten business days of the order",
  czech: "Dodavatel se zavazuje dodat zboží do provozovny kupujícího, přechází ůčinností",
  polish: "Dostawca zobowiązuje się dostarczyć towar w ciągu dziesięciu dni, łódź",
  greek: "Ο Προμηθευτής υποχρεούται να παραδώσει τα ἀγαθά εντός δέκα ημερών",
  cyrillic: "Поставщик обязуется доставить товары в течение десяти дней, Ґрунт ѣ",
} as const satisfies Record<Script, string>;

type CodePointRange = readonly [number, number];

const LATIN: readonly CodePointRange[] = [
  [0x00_00, 0x00_ff],
  [0x2000, 0x206f],
];
const LATIN_EXT: readonly CodePointRange[] = [
  [0x01_00, 0x02_af],
  [0x1e_00, 0x1e_ff],
];
const GREEK: readonly CodePointRange[] = [[0x03_70, 0x03_ff]];
const GREEK_EXT: readonly CodePointRange[] = [[0x1f_00, 0x1f_ff]];
const CYRILLIC: readonly CodePointRange[] = [[0x04_00, 0x04_5f]];
const CYRILLIC_EXT: readonly CodePointRange[] = [[0x04_60, 0x05_2f]];

/** fontsource's subsets for the faces a document's stacks name. */
export const BUNDLED_SUBSETS = {
  Carlito: {
    latin: LATIN,
    "latin-ext": LATIN_EXT,
    greek: GREEK,
    "greek-ext": GREEK_EXT,
    cyrillic: CYRILLIC,
    "cyrillic-ext": CYRILLIC_EXT,
  },
  Lato: { latin: LATIN, "latin-ext": LATIN_EXT },
  Tinos: { latin: LATIN, "latin-ext": LATIN_EXT, greek: GREEK, cyrillic: CYRILLIC },
  // A host UI face no document stack names: its loads must not relay out.
  "Host UI": { latin: LATIN, "latin-ext": LATIN_EXT, greek: GREEK, cyrillic: CYRILLIC },
} as const satisfies Record<string, Record<string, readonly CodePointRange[]>>;

export const HOST_UI_FAMILY = "Host UI";

const FACE_STYLES = ["normal", "italic"] as const;
const FACE_WEIGHTS = [400, 700] as const;

type FaceStyle = (typeof FACE_STYLES)[number];
type FaceWeight = (typeof FACE_WEIGHTS)[number];

type LoadStatus = LayoutFontSetFace["status"];

export class ScriptedFontFace implements LayoutFontSetFace {
  status: LoadStatus = "unloaded";
  readonly waiters: (() => void)[] = [];

  constructor(
    readonly family: string,
    readonly subset: string,
    readonly ranges: readonly CodePointRange[],
    readonly style: FaceStyle,
    readonly weight: FaceWeight,
  ) {}

  covers(codePoint: number): boolean {
    return this.ranges.some(([from, to]) => codePoint >= from && codePoint <= to);
  }

  get key(): string {
    return `${this.family} ${this.style} ${this.weight} ${this.subset}`;
  }
}

type LoadingDoneListener = (event: { fontfaces: readonly LayoutFontSetFace[] }) => void;

type ParsedFont = { style: FaceStyle; weight: FaceWeight; families: string[] };

const FONT_SIZE_TOKEN = /(?:^|\s)\d+(?:\.\d+)?px\s/u;

/** Split a CSS `font` shorthand (as folio builds it) into descriptors and stack. */
const parseFont = (font: string): ParsedFont => {
  const match = FONT_SIZE_TOKEN.exec(font);
  if (!match) {
    throw new Error(`Unparseable font: ${font}`);
  }
  const descriptors = font.slice(0, match.index).split(/\s+/u);
  return {
    style: descriptors.includes("italic") ? "italic" : "normal",
    weight: descriptors.includes("700") || descriptors.includes("bold") ? 700 : 400,
    families: parseFontFamilyList(font.slice(match.index + match[0].length)),
  };
};

const sameFamily = (a: string, b: string): boolean => a.toLowerCase() === b.toLowerCase();

export class ScriptedFontSet implements LayoutFontSet {
  readonly faces: ScriptedFontFace[] = [];
  private readonly listeners = new Set<LoadingDoneListener>();
  private readonly loadingBatch: ScriptedFontFace[] = [];
  /** Queued `loadingdone` batches, dispatched by {@link deliverEvent}. */
  readonly queuedEvents: ScriptedFontFace[][] = [];
  private readyWaiters: (() => void)[] = [];

  constructor() {
    for (const [family, subsets] of Object.entries(BUNDLED_SUBSETS)) {
      for (const [subset, ranges] of Object.entries(subsets)) {
        for (const style of FACE_STYLES) {
          for (const weight of FACE_WEIGHTS) {
            this.faces.push(new ScriptedFontFace(family, subset, ranges, style, weight));
          }
        }
      }
    }
  }

  [Symbol.iterator](): Iterator<LayoutFontSetFace> {
    return this.faces[Symbol.iterator]();
  }

  get ready(): Promise<void> {
    if (this.pending().length === 0) {
      return Promise.resolve();
    }
    return new Promise((resolve) => {
      this.readyWaiters.push(resolve);
    });
  }

  addEventListener(_type: "loadingdone", listener: LoadingDoneListener): void {
    this.listeners.add(listener);
  }

  removeEventListener(_type: "loadingdone", listener: LoadingDoneListener): void {
    this.listeners.delete(listener);
  }

  load(font: string, text = " "): Promise<unknown> {
    const { style, weight, families } = parseFont(font);
    const codePoints = Array.from(text, (character) => character.codePointAt(0) ?? 0);
    const matched = this.faces.filter(
      (face) =>
        face.style === style &&
        face.weight === weight &&
        families.some((family) => sameFamily(family, face.family)) &&
        codePoints.some((codePoint) => face.covers(codePoint)),
    );
    return Promise.all(matched.map((face) => this.request(face)));
  }

  /** What painting `text` in `font` makes the browser fetch. */
  requestForText(font: string, text: string): void {
    void this.load(font, text);
  }

  /** Start loading a face, as a stylesheet rule the page first uses would. */
  request(face: ScriptedFontFace): Promise<void> {
    if (face.status === "loaded") {
      return Promise.resolve();
    }
    if (face.status === "unloaded") {
      face.status = "loading";
      this.loadingBatch.push(face);
    }
    return new Promise((resolve) => {
      face.waiters.push(resolve);
    });
  }

  pending(): ScriptedFontFace[] {
    return this.faces.filter((face) => face.status === "loading");
  }

  /**
   * Finish one load: the face is usable at once, its `load` promises settle,
   * and once nothing is loading `ready` settles and a `loadingdone` is queued.
   */
  complete(face: ScriptedFontFace): void {
    face.status = "loaded";
    for (const resolve of face.waiters.splice(0)) {
      resolve();
    }
    if (this.pending().length > 0) {
      return;
    }
    this.queuedEvents.push(this.loadingBatch.splice(0));
    for (const resolve of this.readyWaiters.splice(0)) {
      resolve();
    }
  }

  deliverEvent(): void {
    const fontfaces = this.queuedEvents.shift();
    if (!fontfaces) {
      return;
    }
    for (const listener of this.listeners) {
      listener({ fontfaces });
    }
  }

  /** Every face loaded before anything measures: the reference font set. */
  loadAll(): void {
    for (const face of this.faces) {
      face.status = "loaded";
    }
  }

  /** The advance of one character in `font`, as this font set draws it now. */
  advance(font: string, character: string): number {
    const { style, weight, families } = parseFont(font);
    const codePoint = character.codePointAt(0) ?? 0;
    for (const family of families) {
      const face = this.faces.find(
        (candidate) =>
          candidate.status === "loaded" &&
          candidate.style === style &&
          candidate.weight === weight &&
          sameFamily(candidate.family, family) &&
          candidate.covers(codePoint),
      );
      if (face) {
        return faceAdvance(face, codePoint);
      }
    }
    // A system fallback, deliberately wider than any bundled face, so a line
    // measured in it breaks elsewhere.
    return 11 + (codePoint % 2);
  }
}

/** Deterministic per-face advances: families and weights differ. */
const faceAdvance = (face: ScriptedFontFace, codePoint: number): number =>
  (face.family.length % 3) + (face.weight === 700 ? 1 : 0) + 5 + (codePoint % 3);

/**
 * A `document` whose canvas measures through `fontSet` and whose `fonts` is it,
 * for the canvas measure backend and `getDocumentFontSet`.
 */
export const scriptedDocument = (fontSet: ScriptedFontSet): object => ({
  fonts: fontSet,
  createElement: () => ({
    getContext: () => ({
      font: "",
      fontKerning: "auto",
      measureText(this: { font: string }, text: string) {
        let width = 0;
        for (const character of text) {
          width += fontSet.advance(this.font, character);
        }
        return { width, actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2 };
      },
    }),
  }),
});
