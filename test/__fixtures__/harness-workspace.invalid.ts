// Deliberate violations of folio-harness-workspaces/no-undeclared-workspace-runtime-import.
import { paragraphLogicalText } from "@stll/docx-core/ops";
import { type Document, paragraphLogicalText as exportedText } from "@stll/docx-core/ops";
export { paragraphLogicalText as reexportedText } from "@stll/docx-core/ops";
export * from "@stll/docx-core/ops";
void import("@stll/docx-core/ops");
void require("@stll/docx-core/ops");

void paragraphLogicalText;
void exportedText;
