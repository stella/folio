# @stll/utf8-boundaries

Shared Unicode boundaries for ZIP entries, streamed text, and previews.

```ts
import {
  createTextReader,
  truncateGraphemes,
  writeZipEntry,
} from "@stll/utf8-boundaries";

writeZipEntry(zip, "word/document.xml", xml);
const reader = createTextReader({ fatal: true, ignoreBOM: true });
const parts = chunks.map((chunk) => reader.read(chunk));
const text = parts.join("") + reader.finish();
const preview = truncateGraphemes(text, 100);
```

`writeZipEntry` UTF-8 encodes complete strings before JSZip receives them. Byte
arrays, buffers, and promised byte arrays retain their byte values. Entry options
are preserved except `binary` and `base64`, which always describe byte input.
JSZip is a peer dependency; the runtime imports only its types.

`createTextReader` keeps one decoder across reads. Call `finish` after the last
chunk; `finishPrefix` deliberately omits an incomplete trailing scalar for a
bounded preview. `truncateCodePoints` counts Unicode scalars; `truncateGraphemes`
uses `Intl.Segmenter` to retain complete extended grapheme clusters.

The `./oxlint` entry exports a default plugin with `no-string-zip-entry` and
`no-chunked-text-codec`. Configure the ZIP rule with `{ ownerPath: "path/to/owner.ts" }`.
The `./testing` entry exports fast-check arbitraries for astral, combining, and RTL
text around stream, 16 KiB, and 64 KiB boundaries, a complete boundary-case census,
and `assertTextRoundTrip` for actual serializer/readback pairs. Fast-check is a peer
needed only by this entry.
