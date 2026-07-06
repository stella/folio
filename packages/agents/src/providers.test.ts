import { describe, expect, test } from "bun:test";

import { toAnthropicTools, toOpenAITools } from "./providers";
import { FOLIO_AGENT_TOOLS, getFolioToolDefinitions } from "./tools";

describe("getFolioToolDefinitions", () => {
  test("returns the same tool list as FOLIO_AGENT_TOOLS", () => {
    expect(getFolioToolDefinitions()).toBe(FOLIO_AGENT_TOOLS);
  });

  test("every tool's inputSchema is a valid object schema", () => {
    for (const definition of FOLIO_AGENT_TOOLS) {
      expect(definition.inputSchema["type"]).toBe("object");
      expect(typeof definition.inputSchema["properties"]).toBe("object");
      expect(definition.inputSchema["additionalProperties"]).toBe(false);
      expect(Array.isArray(definition.inputSchema["required"])).toBe(true);
    }
  });

  test("tool names are unique", () => {
    const names = FOLIO_AGENT_TOOLS.map((definition) => definition.name);
    expect(new Set(names).size).toBe(names.length);
  });
});

describe("toAnthropicTools", () => {
  test("maps name/description/inputSchema onto Anthropic's input_schema shape", () => {
    const anthropicTools = toAnthropicTools(FOLIO_AGENT_TOOLS);
    expect(anthropicTools).toHaveLength(FOLIO_AGENT_TOOLS.length);
    for (const [index, tool] of anthropicTools.entries()) {
      const source = FOLIO_AGENT_TOOLS[index];
      expect(tool.name).toBe(source?.name);
      expect(tool.description).toBe(source?.description);
      expect(tool.input_schema).toBe(source?.inputSchema);
    }
  });
});

describe("toOpenAITools", () => {
  test("maps onto OpenAI's { type: 'function', function } shape", () => {
    const openAiTools = toOpenAITools(FOLIO_AGENT_TOOLS);
    expect(openAiTools).toHaveLength(FOLIO_AGENT_TOOLS.length);
    for (const [index, tool] of openAiTools.entries()) {
      const source = FOLIO_AGENT_TOOLS[index];
      expect(tool.type).toBe("function");
      expect(tool.function.name).toBe(source?.name);
      expect(tool.function.description).toBe(source?.description);
      expect(tool.function.parameters).toBe(source?.inputSchema);
    }
  });
});
