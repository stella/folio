import type {
  FolioAITextRangeHandle,
  FolioDocumentSectionHandle,
  FolioDocumentStory,
  FolioDocumentStoryHandle,
} from "@stll/folio-core/server";

import { FOLIO_AGENT_TOOL_NAMES } from "./types";
import type {
  FolioAgentApplyOperationsSummary,
  FolioAgentBlock,
  FolioAgentChange,
  FolioAgentComment,
  FolioAgentCommentFilter,
  FolioAgentDocumentOutline,
  FolioAgentScopedFindTextResult,
  FolioAgentSectionRead,
  FolioAgentToolName,
  FolioAgentTypedToolDefinition,
  FolioToolCallResult,
} from "./types";

type EmptyArgs = Readonly<Record<string, never>>;

export type FolioAgentFindTextScope =
  | { type: "document" }
  | { type: "section"; handle: FolioDocumentSectionHandle }
  | { type: "page"; page: number }
  | { type: "story"; handle: FolioDocumentStoryHandle };

export type FolioAgentShowInDocumentArgs =
  | { blockId: string; range?: never }
  | { blockId?: never; range: FolioAITextRangeHandle };

export type FolioAgentToolInputByName = {
  [FOLIO_AGENT_TOOL_NAMES.readDocument]: EmptyArgs;
  [FOLIO_AGENT_TOOL_NAMES.getDocumentOutline]: { maxDepth?: number };
  [FOLIO_AGENT_TOOL_NAMES.readSection]: {
    handle: FolioDocumentSectionHandle;
    maxBlocks?: number;
    afterBlockId?: string;
  };
  [FOLIO_AGENT_TOOL_NAMES.listStories]: EmptyArgs;
  [FOLIO_AGENT_TOOL_NAMES.readStory]: { handle: FolioDocumentStoryHandle };
  [FOLIO_AGENT_TOOL_NAMES.findText]: {
    query: string;
    matchCase?: boolean;
    wholeWord?: boolean;
    scope?: FolioAgentFindTextScope;
  };
  [FOLIO_AGENT_TOOL_NAMES.readComments]: { filter?: FolioAgentCommentFilter };
  [FOLIO_AGENT_TOOL_NAMES.readChanges]: EmptyArgs;
  [FOLIO_AGENT_TOOL_NAMES.addComment]: {
    blockId: string;
    quote?: string;
    text: string;
    precondition?: { blockTextHash: string };
  };
  [FOLIO_AGENT_TOOL_NAMES.suggestChanges]: {
    operations: unknown[];
    /** Present when the tool was defined with a `documentVersion` option. */
    documentVersion?: string;
  };
  [FOLIO_AGENT_TOOL_NAMES.replyComment]: {
    commentId: string;
    text: string;
  };
  [FOLIO_AGENT_TOOL_NAMES.resolveComment]: {
    commentId: string;
    reopen?: boolean;
  };
  [FOLIO_AGENT_TOOL_NAMES.readPage]: { page: number };
  [FOLIO_AGENT_TOOL_NAMES.readSelection]: EmptyArgs;
  [FOLIO_AGENT_TOOL_NAMES.scrollToBlock]: { blockId: string };
  [FOLIO_AGENT_TOOL_NAMES.showInDocument]: FolioAgentShowInDocumentArgs;
};

export type FolioAgentToolOutputByName = {
  [FOLIO_AGENT_TOOL_NAMES.readDocument]: FolioAgentBlock[];
  [FOLIO_AGENT_TOOL_NAMES.getDocumentOutline]: FolioAgentDocumentOutline;
  [FOLIO_AGENT_TOOL_NAMES.readSection]: FolioAgentSectionRead;
  [FOLIO_AGENT_TOOL_NAMES.listStories]: FolioDocumentStory[];
  [FOLIO_AGENT_TOOL_NAMES.readStory]: FolioDocumentStory;
  [FOLIO_AGENT_TOOL_NAMES.findText]: FolioAgentScopedFindTextResult;
  [FOLIO_AGENT_TOOL_NAMES.readComments]: FolioAgentComment[];
  [FOLIO_AGENT_TOOL_NAMES.readChanges]: FolioAgentChange[];
  [FOLIO_AGENT_TOOL_NAMES.addComment]: FolioAgentApplyOperationsSummary;
  [FOLIO_AGENT_TOOL_NAMES.suggestChanges]: FolioAgentApplyOperationsSummary;
  [FOLIO_AGENT_TOOL_NAMES.replyComment]: { replied: true };
  [FOLIO_AGENT_TOOL_NAMES.resolveComment]: { resolved: boolean };
  [FOLIO_AGENT_TOOL_NAMES.readPage]: { page: number; totalPages?: number; text: string };
  [FOLIO_AGENT_TOOL_NAMES.readSelection]: { text: string };
  [FOLIO_AGENT_TOOL_NAMES.scrollToBlock]: { scrolled: boolean };
  [FOLIO_AGENT_TOOL_NAMES.showInDocument]: { shown: boolean };
};

export type FolioToolCallResultFor<TName extends FolioAgentToolName> = FolioToolCallResult<
  FolioAgentToolOutputByName[TName]
>;

export const defineFolioAgentToolDefinition = <TName extends FolioAgentToolName>(
  definition: FolioAgentTypedToolDefinition<
    TName,
    FolioAgentToolInputByName[TName],
    FolioAgentToolOutputByName[TName]
  >,
): FolioAgentTypedToolDefinition<
  TName,
  FolioAgentToolInputByName[TName],
  FolioAgentToolOutputByName[TName]
> => definition;
