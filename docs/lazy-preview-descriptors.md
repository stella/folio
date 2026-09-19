# Lazy preview descriptors

Design note. Nothing here is implemented.

## The cost

A `w:drawing` whose graphic is a diagram carries no `a:blip`, so there is no
image to resolve. folio rasterises a placeholder at parse time
(`packages/core/src/docx/diagramPreview.ts`) and stores it as a base64 PNG in
`Image.src`, because `Image.src` is the only way the model can say what a
drawing looks like. The raster is bounded (`MAX_PREVIEW_PIXELS` 1,440,000,
`MAX_PREVIEW_SHAPES` 128) but not small: 7.33 MB per diagram, as base64, held
for the life of the document. The public corpus has 52 packages with diagrams
and reaches 51 MB of preview in one package. The bytes are paid whether or not
anything ever paints the drawing.

Making the raster lazy means the model must be able to say "this drawing looks
like a diagram of these shapes" without saying "here are its pixels". That is a
descriptor, and a descriptor has to travel every path `src` travels today.

## The descriptor

```ts
/** How a drawing with no image data is drawn, without drawing it. */
export type PreviewDescriptor = {
  kind: "diagram";
  /** Shapes already clipped to `MAX_PREVIEW_SHAPES`, in paint order. */
  shapes: readonly PreviewShape[];
  /** The raster's dimensions, so a backend can size without rasterising. */
  pixelWidth: number;
  pixelHeight: number;
};
```

Kept as a discriminated union on `kind` rather than a diagram-shaped record:
chart frames and OLE previews are the same problem and will want the same
field, and a `kind` now is cheaper than migrating a bare `diagram` later.

`shapes` is the intermediate the rasteriser already computes before it paints;
it is small (hundreds of bytes per shape against megabytes of raster) and it
is what every backend needs. Holding it is not a cache: it is the parse
result, and the raster becomes the thing derived on demand.

## What changes

### `@stll/docx-core`: `Image`

`Image` (`packages/docx-core/src/model/content.ts`) gains

```ts
/** Present when the drawing has no image data and is drawn from a description. */
preview?: PreviewDescriptor;
```

`src` keeps its meaning — resolved image data — and stops being set for
diagrams. The two are mutually exclusive in practice; they are not modelled as
a union because `Image` has thirty other fields and a union over the whole
record would churn every consumer. A drawing with neither is the existing
"nothing to paint" case and stays valid.

### Display list: `ImageTable`

`ImageTable.intern` (`packages/core/src/display-list/build/imagePrimitives.ts`)
takes a `src: string` and decodes base64 PNG or JPEG data URLs; a `blob:` URL
already fails with "image bytes are not inline (blob: URLs need a fetch the
builder cannot do)", because the builder does no I/O. A descriptor is the same
shape of problem with the opposite answer: the bytes are not there, but they
can be produced without I/O.

`intern` gains a sibling that takes a descriptor and returns the same
`DisplayImageRef`, so `DisplayList.images` stays
`readonly DisplayImageSource[]` and `DisplayImageSource` keeps its four fields
(`format`, `bytes`, `pixelWidth`, `pixelHeight`). Where the rasterisation
happens is the decision that matters:

- **Eager at intern.** Smallest diff, and it keeps every painter unchanged,
  but it pays the raster for any drawing on a laid-out page whether or not the
  page is visible. It removes the parse-time cost and nothing else.
- **Lazy in `DisplayImageSource`.** `bytes` becomes a thunk the painters call.
  It defers the raster to the paint, but it makes the display list impure and
  no longer a plain serializable value, which the DOM renderer, the PDF
  exporter and the paint-equivalence tests all rely on.
- **Per backend.** `DisplayImageSource` gains a `descriptor` variant and each
  backend draws it natively: the DOM renderer emits shapes, the PDF exporter
  (`packages/core/src/pdf/images.ts`) draws vector operators. No raster is ever
  produced. Largest diff, best result, and it is the only option that also
  improves PDF output, which today embeds a bitmap of a vector drawing.

The third is the one to aim at; the first is a legitimate intermediate step
that can ship first and be narrowed later, because it does not change any
contract.

### ProseMirror

`toProseDoc` copies `image.src || ""` into the node's `src` attr
(`packages/core/src/prosemirror/conversion/toProseDoc.ts:3744`). A node whose
`src` is `""` is an image the editor cannot draw, so the attr has to carry the
descriptor or a handle to it.

The attr cannot hold the descriptor itself. PM attrs are serialized into
persisted editor state and into collaboration updates; putting shape geometry
there moves the cost rather than removing it, and multiplies it by every
client in a collaboration session. The attr should hold a stable id, and the
descriptor should live beside the document where the conversion can reach it.

This is the load-bearing decision of the whole change, and it is where the
persisted-state compatibility question sits: an existing persisted document has
a `src` holding a base64 raster. Reading one must keep working. Since the attr
gains a field rather than changing one, an old state parses with the new
schema and paints from `src` exactly as before; a new state has an empty `src`
and a preview id. No dual-write is needed, and no migration of stored state —
the boundary is "documents saved before this lands keep their raster", and it
closes when those documents are next re-parsed from their `.docx`.

### Adapters and parity

`@stll/folio-react` and `@stll/folio-vue` do not touch image bytes and need no
change, but `bun run check:parity-contract` and `bun run check:export-parity`
run on any public-contract change and must be green.

## api-report impact

- `api-reports/docx-core/*.api.md`: `Image.preview` and the exported
  `PreviewDescriptor` / `PreviewShape` types.
- `api-reports/core/*.api.md`: whatever the `ImageTable` and display-list
  changes expose. `DisplayImageSource` is public; a new variant is an additive
  change to it.

Additive throughout, so a minor bump for `@stll/docx-core` and
`@stll/folio-core`.

## Migration order

1. `PreviewDescriptor` in `@stll/docx-core`, and the parser filling it
   alongside the raster it already produces. Nothing reads it; nothing changes.
2. `ImageTable` interning a descriptor by rasterising at intern. The parser
   stops writing `src` for diagrams. Parse-time memory drops here, and this is
   the step the corpus numbers above measure.
3. The PM attr id and the side table, so the editor paints diagrams again.
4. Per-backend drawing in the DOM renderer and the PDF exporter, retiring the
   raster.

Steps 1 and 2 are where the 7.33 MB per diagram goes; 3 restores editor
parity; 4 is the quality win. Each is independently shippable.

## Expected saving

Per package, `preview raster bytes - descriptor bytes`. At 7.33 MB of base64
raster per diagram against shapes bounded by `MAX_PREVIEW_SHAPES` (128), the
descriptor is three orders of magnitude smaller, so the saving is effectively
the whole of it: 7.33 MB per diagram, up to 51 MB on the heaviest package in
the corpus, across the 52 packages that have diagrams at all. After step 4 the
raster is never built, so the peak during parse falls by the same amount rather
than merely being freed sooner.
