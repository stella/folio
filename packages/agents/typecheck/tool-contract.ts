import type { FolioAgentBridge } from "../src/bridge";
import { executeFolioToolCall } from "../src/execute";
import { FOLIO_AGENT_TOOL_NAMES } from "../src/types";
import type {
  FolioAgentBlock,
  FolioAgentComment,
  FolioAgentScopedFindTextResult,
} from "../src/index";
import type { FolioToolCallResultFor } from "../src/tool-contract";

declare const bridge: FolioAgentBridge;

// @ts-expect-error empty-input tools reject undeclared fields
executeFolioToolCall(FOLIO_AGENT_TOOL_NAMES.readDocument, { unexpected: true }, bridge);

const readDocumentResult = executeFolioToolCall(FOLIO_AGENT_TOOL_NAMES.readDocument, {}, bridge);
const readDocumentCheck: FolioToolCallResultFor<typeof FOLIO_AGENT_TOOL_NAMES.readDocument> =
  readDocumentResult;
if (readDocumentCheck.ok) {
  const blocks: FolioAgentBlock[] = readDocumentCheck.result;
  void blocks;
}

const findTextResult = executeFolioToolCall(
  FOLIO_AGENT_TOOL_NAMES.findText,
  { query: "Heading" },
  bridge,
);
if (findTextResult.ok) {
  const matches: FolioAgentScopedFindTextResult = findTextResult.result;
  void matches;
}

const readCommentsResult = executeFolioToolCall(FOLIO_AGENT_TOOL_NAMES.readComments, {}, bridge);
if (readCommentsResult.ok) {
  const comments: FolioAgentComment[] = readCommentsResult.result;
  void comments;
}

const replyCommentResult = executeFolioToolCall(
  FOLIO_AGENT_TOOL_NAMES.replyComment,
  { commentId: "17", text: "Clarifying now." },
  bridge,
);
if (replyCommentResult.ok) {
  const replied: true = replyCommentResult.result.replied;
  void replied;
}
