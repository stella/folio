// Shapes folio-harness-workspaces/no-undeclared-workspace-runtime-import must accept.
import type { FolioAgentToolName } from "@stll/folio-agents";
import { type FolioAgentVersionDiff } from "@stll/folio-agents/compare";
import { createEmptyDocument } from "@stll/folio-core/utils/createDocument";
import { paragraphLogicalText } from "../../packages/docx-core/src/ops/offsets";

export type { FolioAgentVersionDiff } from "@stll/folio-agents/compare";

void import("@stll/folio-core/controller/folioEditor");

void require("@stll/folio-core");

declare const importedToolName: FolioAgentToolName;
declare const importedDiff: FolioAgentVersionDiff;
void importedToolName;
void importedDiff;
void createEmptyDocument;
void paragraphLogicalText;
