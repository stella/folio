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
 */

/** What identifies a preview in the model, and how much of it a package may keep. */
type PreviewKind = {
  readonly mimeType: string;
  readonly filename: string;
  readonly srcPrefix: string;
  readonly maxPackageCharacters: number;
};

export const VML_PREVIEW_DATA_URL_PREFIX = "data:image/svg+xml;charset=utf-8,";

const MEBIBYTE = 1024 * 1024;

export const PREVIEW_KINDS = {
  /** A VML shape folio renders rather than projects (`v:shape`, `v:rect`, ...). */
  vmlShape: {
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
   * contained. It is a `PreviewDescriptor` now, which both backends draw.
   *
   * What the producer still takes from this entry is the mime type and the
   * filename it stamps on the image. Its `srcPrefix` and character cap match
   * nothing, because a diagram no longer has a `src` for the budget to charge.
   */
  smartArt: {
    mimeType: "image/png",
    filename: "smartart-preview.png",
    srcPrefix: "data:image/png;base64,",
    maxPackageCharacters: 64 * MEBIBYTE,
  },
} as const satisfies Record<string, PreviewKind>;

type PreviewKindName = keyof typeof PREVIEW_KINDS;

const KIND_NAMES = Object.keys(PREVIEW_KINDS) as PreviewKindName[];

/**
 * The kind a model image was generated as, or `undefined` for one the package
 * actually carries. A generated preview has no relationship behind it, so
 * `rId` is empty; the filename and data-URL prefix name which producer made it.
 */
const previewKindOf = (value: object): PreviewKindName | undefined => {
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
  return KIND_NAMES.find((name) => {
    const kind = PREVIEW_KINDS[name];
    return (
      mimeType === kind.mimeType && filename === kind.filename && src.startsWith(kind.srcPrefix)
    );
  });
};

/** Per-kind allowances for one package, defaulting to the table's caps. */
export type PreviewBudgetOverrides = Partial<Record<PreviewKindName, number>>;

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
  const remaining = new Map<PreviewKindName, number>(
    KIND_NAMES.map((name) => [
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
      // an entry for every kind name it can return.
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
