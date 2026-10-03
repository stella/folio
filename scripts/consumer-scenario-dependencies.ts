import manifest from "../package.json";
import docxCoreManifest from "../packages/docx-core/package.json";

/** Third-party packages installed and accepted by the isolated consumer guard. */
export const CONSUMER_DEPENDENCIES = [
  "prosemirror-state@^1.4.4",
  "prosemirror-model@^1.25.9",
  `better-result@${manifest.devDependencies["better-result"]}`,
  `marked@${docxCoreManifest.dependencies.marked}`,
  `fast-xml-parser@${docxCoreManifest.dependencies["fast-xml-parser"]}`,
];
