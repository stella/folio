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
 * Producer and budget have to agree on what a preview looks like. They used to
 * agree by coincidence: the VML producer wrote its filename as a literal and
 * the budget recognized it as a constant in another file, so renaming one
 * would have stopped the other charging for it without anything failing. The
 * table below is that agreement written once, and a producer builds its image
 * out of the entry the budget matches against.
 *
 * An entry says what backs the preview, because that decides what there is to
 * agree on. A source-backed preview is image data in `src`, recognized by the
 * strings its producer stamps on it and charged by the character. A
 * descriptor-backed preview is no picture at all, so it carries none of those
 * strings: a table that offered them would be offering a cap over nothing and
 * a prefix nothing can start with.
 */

/**
 * A preview the model carries as image data in `src`.
 *
 * The three strings are how the budget recognizes what a producer made, and
 * the cap is how much of that data one package may keep.
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
 * the character budget to recognize or to charge, and no entry of its own to
 * drift out of agreement with its producer. What the entry still says is that
 * folio attaches this preview and that the budget passes over it deliberately,
 * which the switch below has to keep answering as kinds are added.
 */
type DescriptorBackedPreviewKind = {
  readonly backing: "descriptor";
};

/** What backs a preview in the model, and how much of it a package may keep. */
type PreviewKind = SourceBackedPreviewKind | DescriptorBackedPreviewKind;

type PreviewKindName = "vmlShape" | "smartArt";

export const VML_PREVIEW_DATA_URL_PREFIX = "data:image/svg+xml;charset=utf-8,";

const MEBIBYTE = 1024 * 1024;

export const PREVIEW_KINDS = {
  /** A VML shape folio renders rather than projects (`v:shape`, `v:rect`, ...). */
  vmlShape: {
    backing: "source",
    mimeType: "image/svg+xml",
    filename: "vml-shape-preview.svg",
    srcPrefix: VML_PREVIEW_DATA_URL_PREFIX,
    maxPackageCharacters: 8 * MEBIBYTE,
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

const KIND_NAMES = Object.keys(PREVIEW_KINDS) as PreviewKindName[];

/**
 * Whether a package charges this kind, narrowing the name to a source-backed
 * one: the predicate reads the same field `SourceBackedKindName` selects on.
 */
const isSourceBacked = (name: PreviewKindName): name is SourceBackedKindName => {
  const kind = PREVIEW_KINDS[name];
  switch (kind.backing) {
    case "source":
      return true;
    case "descriptor":
      return false;
    default:
      return kind satisfies never;
  }
};

const SOURCE_BACKED_KIND_NAMES = KIND_NAMES.filter(isSourceBacked);

/**
 * The kind a model image was generated as, or `undefined` for one the package
 * actually carries. A generated preview has no relationship behind it, so
 * `rId` is empty; the filename and data-URL prefix name which producer made it.
 *
 * Only a source-backed kind can be named this way. A descriptor-backed preview
 * has no `src` to match, so no image the model carries can be charged to it.
 */
const previewKindOf = (value: object): SourceBackedKindName | undefined => {
  if (
    !("type" in value) ||
    value.type !== "image" ||
    !("rId" in value) ||
    value.rId !== "" ||
    !("src" in value) ||
    typeof value.src !== "string" ||
    !("mimeType" in value) ||
    !("filename" in value)
  ) {
    return undefined;
  }
  const { src, mimeType, filename } = value;
  return SOURCE_BACKED_KIND_NAMES.find((name) => {
    const kind = PREVIEW_KINDS[name];
    return (
      mimeType === kind.mimeType && filename === kind.filename && src.startsWith(kind.srcPrefix)
    );
  });
};

/**
 * Per-kind allowances for one package, defaulting to the table's caps.
 *
 * Only a source-backed kind has an allowance: there is no number that would
 * change what the budget does to a descriptor.
 */
export type PreviewBudgetOverrides = Partial<Record<SourceBackedKindName, number>>;

/**
 * Charge every generated preview in the model against its kind's allowance and
 * drop the `src` of those past it. Dropping leaves the image in place with its
 * size and wrap, so the page still reserves the space the drawing occupies,
 * and never touches the preserved XML the package saves from.
 */
export const enforcePackagePreviewBudget = (
  root: unknown,
  overrides: PreviewBudgetOverrides = {},
): void => {
  const remaining = new Map<SourceBackedKindName, number>(
    SOURCE_BACKED_KIND_NAMES.map((name) => [
      name,
      Math.max(0, overrides[name] ?? PREVIEW_KINDS[name].maxPackageCharacters),
    ]),
  );
  const visited = new WeakSet<object>();

  const visit = (value: unknown): void => {
    if (value === null || typeof value !== "object" || visited.has(value)) {
      return;
    }
    visited.add(value);
    if (value instanceof ArrayBuffer || ArrayBuffer.isView(value)) {
      return;
    }
    if (value instanceof Map) {
      for (const child of value.values()) {
        visit(child);
      }
      return;
    }
    if (Array.isArray(value)) {
      for (const child of value) {
        visit(child);
      }
      return;
    }
    const kind = previewKindOf(value);
    if (kind !== undefined) {
      // SAFETY: `previewKindOf` matched on a string `src`, and `remaining` has
      // an entry for every source-backed name it can return.
      const image = value as { src?: string };
      const length = (image.src as string).length;
      const left = remaining.get(kind) as number;
      if (length <= left) {
        remaining.set(kind, left - length);
      } else {
        remaining.set(kind, 0);
        delete image.src;
      }
    }
    for (const child of Object.values(value)) {
      visit(child);
    }
  };

  visit(root);
};
