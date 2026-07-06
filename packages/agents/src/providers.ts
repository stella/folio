import type { FolioAgentToolDefinition } from "./types";

/** Anthropic Messages API tool-definition shape (`input_schema`, snake_case). */
export type AnthropicToolDefinition = {
  name: string;
  description: string;
  input_schema: Record<string, unknown>;
};

/** Map folio's provider-neutral tool definitions onto Anthropic's tool shape. */
export const toAnthropicTools = (
  definitions: FolioAgentToolDefinition[],
): AnthropicToolDefinition[] =>
  definitions.map((definition) => ({
    name: definition.name,
    description: definition.description,
    input_schema: definition.inputSchema,
  }));

/** OpenAI Chat Completions / Responses API function-tool shape. */
export type OpenAIToolDefinition = {
  type: "function";
  function: {
    name: string;
    description: string;
    parameters: Record<string, unknown>;
  };
};

/** Map folio's provider-neutral tool definitions onto OpenAI's function-tool shape. */
export const toOpenAITools = (definitions: FolioAgentToolDefinition[]): OpenAIToolDefinition[] =>
  definitions.map((definition) => ({
    type: "function",
    function: {
      name: definition.name,
      description: definition.description,
      parameters: definition.inputSchema,
    },
  }));
