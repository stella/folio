import manifest from "../package.json";

/** Third-party packages installed and accepted by the isolated consumer guard. */
export const CONSUMER_DEPENDENCIES = [
  "prosemirror-state@^1.4.4",
  "prosemirror-model@^1.25.9",
  `better-result@${manifest.devDependencies["better-result"]}`,
];
