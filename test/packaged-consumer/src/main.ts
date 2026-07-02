// Import the surfaces a real consumer wires up, from the PACKED packages. The
// editor pulls the layout engine, which pulls the off-main-thread font-metrics
// worker (`measureWorker` -> `new Worker(new URL(...))`). Importing the
// `measureWorker` subpath directly pins that module into the build graph so the
// production Vite build always resolves the worker URL target (the class of
// regression this consumer exists to catch), independent of how the editor
// tree-shakes. The messages subpath is the bundled UI catalog.
import { DocxEditor, FolioUIProvider } from "@stll/folio-react";
import { getFolioMessages } from "@stll/folio-react/messages";
import { canPrefetchMeasurement } from "@stll/folio-core/layout-engine/measure/measureWorker";

// Reference every import so nothing is eliminated before the bundler's worker
// transform runs. Exported (not executed): the build is what is under test.
export const consumed = {
  DocxEditor,
  FolioUIProvider,
  getFolioMessages,
  canPrefetchMeasurement,
};
