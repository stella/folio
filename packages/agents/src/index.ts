export type { FolioAgentBridge } from "./bridge";
export type {
  CreateEditorRefBridgeOptions,
  FolioAgentEditorApplyDocumentOperationsOptions,
  FolioAgentEditorRefLike,
} from "./bridges/editor-ref";
export { createEditorRefBridge } from "./bridges/editor-ref";
export type { CreateReviewerBridgeOptions } from "./bridges/reviewer";
export { createReviewerBridge } from "./bridges/reviewer";
export type {
  FolioAgentBlockDiff,
  FolioAgentVersionDiff,
  FolioAgentVersionDiffSegment,
} from "./compare";
export { compareDocxVersions, formatVersionDiffForLLM } from "./compare";
export { executeFolioToolCall } from "./execute";
export {
  FOLIO_DOCUMENT_OPERATION_BATCH_JSON_SCHEMA,
  FOLIO_DOCUMENT_OPERATION_JSON_SCHEMA,
  folioDocumentOperationBatchSchema,
} from "./operation-schema";
export type { ParseAddCommentResult, ParseSuggestChangesResult } from "./parse";
export { parseAddCommentInput, parseSuggestChangesInput } from "./parse";
export type { AnthropicToolDefinition, OpenAIToolDefinition } from "./providers";
export { toAnthropicTools, toOpenAITools } from "./providers";
export { FOLIO_AGENT_TOOLS, getFolioToolDefinitions } from "./tools";
export type {
  FolioAgentApplyOperationsSummary,
  FolioAgentBlock,
  FolioAgentChange,
  FolioAgentComment,
  FolioAgentCommentFilter,
  FolioAgentCommentReply,
  FolioAgentTextMatch,
  FolioAgentToolDefinition,
  FolioAgentToolName,
  FolioToolCallResult,
} from "./types";
export { FOLIO_AGENT_TOOL_NAMES } from "./types";
