// Deliberate violations of folio-harness-workspaces/no-undeclared-workspace-runtime-import.
import { compareDocxVersions } from "@stll/folio-agents/compare";
import {
  type FolioAgentVersionDiff,
  compareDocxVersions as exportedCompare,
} from "@stll/folio-agents/compare";

export { compareDocxVersions as reexportedCompare } from "@stll/folio-agents/compare";
export * from "@stll/folio-agents/compare";
void import("@stll/folio-agents/compare");
void require("@stll/folio-agents/compare");

declare const importedDiff: FolioAgentVersionDiff;
void importedDiff;
void compareDocxVersions;
void exportedCompare;
