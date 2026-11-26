---
"@stll/folio-react": minor
"@stll/folio-vue": minor
---

The dialog components are no longer exported from the package root. Import them from the `dialogs` subpath instead:

```ts
// before
import { HyperlinkDialog } from "@stll/folio-react";
// after
import { HyperlinkDialog } from "@stll/folio-react/dialogs";
```

The same applies to `@stll/folio-vue` and `@stll/folio-vue/dialogs`. This lets bundlers keep the editor's dialogs out of the startup chunk: each loads the first time it opens. The Vue editor now loads its dialogs (find and replace, hyperlink, symbol, image and table properties, page setup, numbering value, watermark) on first open and mounts each only while it is open, as the React editor does. A Vue dialog mounted already open sets itself up as one opened later does. In both editors the Insert Symbol dialog keeps its recent symbols across openings.
