# @stll/folio-fonts

The document fonts folio measures and paints with: metric-compatible,
open-licence stand-ins for the families documents commonly name, so a page
lays out the same width whether or not the authored font is installed.

| Family           | Stands in for                                  |
| ---------------- | ---------------------------------------------- |
| Carlito          | Calibri                                        |
| Caladea          | Cambria                                        |
| Arimo            | Arial, and the generic sans-serif              |
| Tinos            | Times New Roman, and the generic serif         |
| Cousine          | Courier New, and the generic monospace         |
| Lato             | Aptos and Aptos Display                        |
| Source Sans 3    | Segoe UI                                       |
| Noto Sans Arabic | Dubai, and Arabic script in every other family |

`@stll/folio-react`, `@stll/folio-vue` and `@stll/folio-cli` depend on this
package and wire it up themselves; install it directly only to use the faces
outside those packages. It has no runtime dependencies.

## In a browser

```css
@import "@stll/folio-fonts/fonts.css";
```

Each face is declared as `.woff2`, cut into the `unicode-range` subsets its
source ships, so the browser downloads a file only when a page paints one of
its code points. The `latin`, `latin-ext` and `arabic` subsets list their
`.woff` after it, the file headless measurement reads, so a test page can
measure the same bytes; a browser that reads `.woff2` never fetches it. Weights 400 and 700, upright and italic (Noto Sans Arabic
has no italic).

## In a bundler or React Native

Every file is exported by its path, so a bundler that cannot read CSS can
import a face directly:

```ts
const carlito = require("@stll/folio-fonts/fonts/carlito/files/carlito-latin-400-normal.woff2");
```

`FOLIO_FONT_FACES` lists every face with its family, weight, style, subset,
`unicode-range` and package-relative paths; `FOLIO_FONT_FILES` lists every
file under `fonts/`.

## In Node.js

Headless layout and PDF export read the `.woff` files of the `latin`,
`latin-ext` and `arabic` subsets:

```ts
import { createFontsourceFaces } from "@stll/folio-core/fonts/fontsourceFaces";
import { folioFontFiles } from "@stll/folio-fonts/node";

const faces = createFontsourceFaces(folioFontFiles);
```

`folioFontFiles.read(directory, path)` reads `fonts/<directory>/<path>`, the
layout of the `@fontsource` package each family comes from, and answers `null`
for anything the package does not ship. When a bundler moves the code away
from the files, copy `fonts/` and pass its location to `createFolioFontFiles`.

## Licences

The font files are licensed under the SIL Open Font License 1.1; each family's
licence is in `licenses/` and beside its files. The package's code and the
generated stylesheet are Apache-2.0. `THIRD_PARTY_NOTICES.md` names each
family, where it comes from and its licence.

The files are copied from the `@fontsource` packages by
`bun run generate:folio-fonts` in the folio repository, which also writes
`fonts.css`, the notices and the catalogue; CI checks them byte for byte.
