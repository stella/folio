import { withCanonicalCommand } from "../../../../canonicalCommands";

export const BoldExtension = {
  schemaMarkName: "bold",
  command: withCanonicalCommand(() => true, () => []),
};
