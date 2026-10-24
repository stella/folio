import { withRunFormattingLock } from "../../../../extensions/marks/markUtils";

export const BoldExtension = {
  schemaMarkName: "bold",
  command: withRunFormattingLock(() => true, () => []),
};
