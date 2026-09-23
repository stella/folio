/**
 * Every synthetic preview a parse attaches to the model, and the budget each
 * one answers to.
 *
 * A preview is not document content. It is a drawing folio makes up so a shape
 * it cannot project still occupies the page, and the package round-trips from
 * its preserved XML whether the preview exists or not. So a preview is the one
 * thing in the model that may be dropped, and a package that would retain more
 * of it than it is worth has it dropped rather than being refused.
 *
 * Producer and budget have to agree on which images are previews. They used to
 * agree by coincidence: the VML producer wrote its filename as a literal and
 * the budget recognized it as a constant in another file, so renaming one
 * would have stopped the other charging for it without anything failing. Then
 * they agreed on the strings in the table below, and the budget found previews
 * by walking the whole model for images that carried them: every object of
 * every parse visited, for the handful of previews a package holds.
 *
 * The agreement is now a ledger the parse owns. The table is private to this
 * module, so the ledger's factory is the only code that can stamp a preview's
 * strings on an image, and the factory registers every image it builds. The
 * budget charges the ledger and nothing else: a preview it has not seen is a
 * preview nothing could have built.
 *
 * An entry says what backs the preview, because that decides what there is to
 * agree on. A source-backed preview is image data in `src`, built by the
 * ledger and charged by the character. A descriptor-backed preview is no
 * picture at all, so it carries none of those strings: a table that offered
 * them would be offering a cap over nothing and a prefix nothing can start
 * with.
 */

import type { Image } from "../types/document";

/**
 * A preview the model carries as image data in `src`.
 *
 * The three strings are what the ledger stamps on the image it builds, and the
 * cap is how much of that data one package may keep.
 */
type SourceBackedPreviewKind = {
  readonly backing: "source";
  readonly mimeType: string;
  readonly filename: string;
  readonly srcPrefix: string;
  readonly maxPackageCharacters: number;
};

/**
 * A preview the model carries as a `PreviewDescriptor`, which a backend draws.
 *
 * There is no image data for it anywhere in the model, so it has nothing for
 * the character budget to charge, and no entry of its own to drift out of
 * agreement with its producer. What the entry still says is that folio
 * attaches this preview and that the budget passes over it deliberately: the
 * ledger's factory accepts only a source-backed kind.
 */
type DescriptorBackedPreviewKind = {
  readonly backing: "descriptor";
};

/** What backs a preview in the model, and how much of it a package may keep. */
type PreviewKind = SourceBackedPreviewKind | DescriptorBackedPreviewKind;

type PreviewKindName = "vmlShape" | "wpGroup" | "smartArt";

/** Both SVG producers encode their drawing into the URL rather than base64. */
const SVG_PREVIEW_DATA_URL_PREFIX = "data:image/svg+xml;charset=utf-8,";

const MEBIBYTE = 1024 * 1024;

/**
 * Eight times what a producer may spend on one drawing, which is a package
 * holding a handful of the largest previews folio will build.
 */
const SVG_PREVIEW_PACKAGE_CHARACTERS = 8 * MEBIBYTE;

const PREVIEW_KINDS = {
  /** A VML shape folio renders rather than projects (`v:shape`, `v:rect`, ...). */
  vmlShape: {
    backing: "source",
    mimeType: "image/svg+xml",
    filename: "vml-shape-preview.svg",
    srcPrefix: SVG_PREVIEW_DATA_URL_PREFIX,
    maxPackageCharacters: SVG_PREVIEW_PACKAGE_CHARACTERS,
  },
  /**
   * A WordprocessingGroup (`wpg:wgp`) folio renders rather than projects.
   *
   * The group's shapes, their text and any picture they carry are drawn into
   * one SVG, so a package with many groups retains as much generated text as a
   * package of VML shapes does, from a producer bounded per drawing the same
   * way. It was once missing from this table, which was not a preview with a
   * generous cap but a preview with none: the budget never recognized it, so
   * no package was ever charged for one.
   */
  wpGroup: {
    backing: "source",
    mimeType: "image/svg+xml",
    filename: "wordprocessing-group.svg",
    srcPrefix: SVG_PREVIEW_DATA_URL_PREFIX,
    maxPackageCharacters: SVG_PREVIEW_PACKAGE_CHARACTERS,
  },
  /**
   * A SmartArt diagram: its extent filled with one flat rectangle per shape.
   *
   * It used to be a raster in `src`, because the display list decoded only
   * base64 PNG and JPEG and a vector preview would have been missing from
   * every backend but the DOM. That cost 7.3 MB of data URL per diagram
   * whatever the package weighed, up to 51.3 MB in one corpus package, because
   * it followed the extent the author chose rather than anything the drawing
   * contained. It is a `PreviewDescriptor` now, which both backends draw, and
   * the bound on it is the shape cap the parse applies to the description.
   */
  smartArt: {
    backing: "descriptor",
  },
} as const satisfies Record<PreviewKindName, PreviewKind>;

/** The kinds a package charges: those whose preview is image data it retains. */
type SourceBackedKindName = {
  [Name in PreviewKindName]: (typeof PREVIEW_KINDS)[Name]["backing"] extends "source"
    ? Name
    : never;
}[PreviewKindName];

/**
 * Per-kind allowances for one package, defaulting to the table's caps.
 *
 * Only a source-backed kind has an allowance: there is no number that would
 * change what the budget does to a descriptor.
 */
export type PreviewBudgetOverrides = Partial<Record<SourceBackedKindName, number>>;

/**
 * Everything a preview image states except what the ledger stamps on it: its
 * geometry, wrap, position and names. A preview has no relationship behind
 * it, so its frame cannot name one.
 */
type PreviewFrame = Omit<Image, "type" | "rId" | "src" | "mimeType" | "filename">;

/** The one way a producer builds a preview image. */
export type PreviewLedger = {
  /**
   * Build a `kind` preview that draws `svg` over `frame`, registered against
   * the package's allowance for that kind.
   */
  readonly svgImage: (kind: SourceBackedKindName, svg: string, frame: PreviewFrame) => Image;
  /**
   * Withdraw every preview inside content the parse built and then discarded,
   * so it spends no allowance. Walks only `discarded`.
   */
  readonly release: (discarded: unknown) => void;
};

/** One parse's ledger, and the budget that charges it. */
export type PackagePreviewBudget = {
  readonly ledger: PreviewLedger;
  /**
   * Charge every registered preview against its kind's allowance and drop the
   * `src` of those past it. Dropping leaves the image in place with its size
   * and wrap, so the page still reserves the space the drawing occupies, and
   * never touches the preserved XML the package saves from.
   *
   * Previews are charged in the order the parse built them, which is the
   * order it read the parts: the body (a text box's content right after the
   * paragraph that anchors it), then headers and footers in relationship
   * order, footnotes, endnotes, and comments last. Past an allowance, the
   * previews built first keep their `src`.
   */
  readonly enforce: (overrides?: PreviewBudgetOverrides) => void;
};

type LedgerEntry = { readonly kind: SourceBackedKindName; readonly image: Image };

/**
 * One parse's previews and their budget.
 *
 * A parse that builds content and then discards it (a note repeating an id, a
 * separator note, a duplicate comment, a watermark's host paragraph) releases
 * that content, or its previews would spend allowance no retained preview can
 * use. A missed release errs towards dropping a preview, never towards
 * retaining more than the cap; the budget's tests pin that a package's ledger
 * holds exactly the previews its model carries.
 */
export const createPackagePreviewBudget = (): PackagePreviewBudget => {
  const entries: LedgerEntry[] = [];
  const registered = new WeakSet<object>();
  const released = new WeakSet<object>();
  const release = (value: unknown, seen: WeakSet<object>): void => {
    if (
      value === null ||
      typeof value !== "object" ||
      seen.has(value) ||
      value instanceof ArrayBuffer ||
      ArrayBuffer.isView(value)
    ) {
      return;
    }
    seen.add(value);
    if (registered.has(value)) {
      released.add(value);
      return;
    }
    for (const child of value instanceof Map ? value.values() : Object.values(value)) {
      release(child, seen);
    }
  };
  return {
    ledger: {
      svgImage: (kind, svg, frame) => {
        const { srcPrefix, mimeType, filename } = PREVIEW_KINDS[kind];
        const image: Image = {
          type: "image",
          src: `${srcPrefix}${encodeURIComponent(svg)}`,
          mimeType,
          filename,
          ...frame,
        };
        entries.push({ kind, image });
        registered.add(image);
        return image;
      },
      release: (discarded) => {
        release(discarded, new WeakSet());
      },
    },
    enforce: (overrides = {}) => {
      const remaining = new Map<SourceBackedKindName, number>();
      for (const { kind, image } of entries) {
        if (image.src === undefined || released.has(image)) {
          continue;
        }
        const left =
          remaining.get(kind) ??
          Math.max(0, overrides[kind] ?? PREVIEW_KINDS[kind].maxPackageCharacters);
        if (image.src.length <= left) {
          remaining.set(kind, left - image.src.length);
        } else {
          remaining.set(kind, 0);
          delete image.src;
        }
      }
    },
  };
};

/**
 * A ledger for a read that builds no package: one part parsed on its own,
 * whose previews no package budget charges.
 */
export const standalonePreviewLedger = (): PreviewLedger => createPackagePreviewBudget().ledger;
