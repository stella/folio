// Shapes folio-harness-workspaces/no-undeclared-workspace-runtime-import must accept.
import type { Document } from "@stll/docx-core";
import { type Paragraph } from "@stll/docx-core/model";
export type { Paragraph } from "@stll/docx-core/model";
import { createEmptyDocument } from "@stll/folio-core/utils/createDocument";
import { paragraphLogicalText } from "../../packages/docx-core/src/ops/offsets";
void import("@stll/folio-core/controller/folioEditor");
void require("@stll/folio-core");
declare const importedDocument: Document;
declare const importedParagraph: Paragraph;
void importedDocument;
void importedParagraph;
void createEmptyDocument;
void paragraphLogicalText;
