# @stll/utf8-boundaries

Shared Unicode boundaries for streamed text and previews.

```ts
import { createTextReader, truncateGraphemes } from "@stll/utf8-boundaries";

const reader = createTextReader({ fatal: true, ignoreBOM: true });
const parts = chunks.map((chunk) => reader.read(chunk));
const text = parts.join("") + reader.finish();
const preview = truncateGraphemes(text, 100);
```

`createTextReader` keeps one decoder across reads. Call `finish` after the last
chunk; `finishPrefix` deliberately omits an incomplete trailing scalar for a
bounded preview. `truncateCodePoints` counts Unicode scalars; `truncateGraphemes`
uses `Intl.Segmenter` to retain complete extended grapheme clusters.

The `./oxlint` entry exports a default plugin with `no-chunked-text-codec`.
The `./testing` entry exports fast-check arbitraries for astral, combining, and RTL
text around stream, 16 KiB, and 64 KiB boundaries, a complete boundary-case census,
and `assertTextRoundTrip` for actual serializer/readback pairs. Fast-check is a peer
needed only by this entry.
