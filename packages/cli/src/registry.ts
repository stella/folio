/**
 * The one tool vocabulary behind both surfaces. Every entry is an MCP tool
 * (its `name`, description, and envelope-wrapped input schema) and one or
 * more `folio` commands whose flags are generated from the same schema. Tools
 * that exist in `@stll/folio-agents` reuse its definitions verbatim; the file
 * envelope (`path`, `fileVersion`, ...) is added here.
 */

import { panic } from "better-result";

import { getFolioToolDefinitions } from "@stll/folio-agents/tools";
import { FOLIO_AGENT_TOOL_NAMES, type FolioAgentToolName } from "@stll/folio-agents/types";

export type JsonObjectSchema = {
  readonly type: "object";
  readonly properties: Readonly<Record<string, unknown>>;
  readonly required: readonly string[];
  readonly additionalProperties: false;
};

/** One `folio <name>` command; `preset` fixes arguments the command implies. */
export type FolioCommandSpec = {
  readonly name: string;
  readonly summary: string;
  readonly preset?: Readonly<Record<string, unknown>>;
};

type FolioToolSpecBase = {
  readonly name: string;
  readonly description: string;
  /** The tool's own arguments, without the file envelope. */
  readonly argsSchema: JsonObjectSchema;
  readonly commands: readonly FolioCommandSpec[];
};

/**
 * `agentRead` runs a read-only `@stll/folio-agents` tool against one file and
 * never writes.
 */
export type FolioFileToolSpec = FolioToolSpecBase & {
  readonly type: "agentRead";
  readonly agentTool: FolioAgentToolName;
};

export type FolioFileToolName = FolioFileToolSpec["name"];

const agentDefinitions = new Map(
  getFolioToolDefinitions().map((definition) => [definition.name, definition]),
);

const isJsonObjectSchema = (schema: unknown): schema is JsonObjectSchema =>
  typeof schema === "object" &&
  schema !== null &&
  "type" in schema &&
  schema.type === "object" &&
  "properties" in schema &&
  typeof schema.properties === "object" &&
  schema.properties !== null &&
  "required" in schema &&
  Array.isArray(schema.required);

const agentDefinition = (name: FolioAgentToolName) => {
  const definition =
    agentDefinitions.get(name) ?? panic(`@stll/folio-agents has no tool named ${name}`);
  const schema = definition.inputSchema;
  if (!isJsonObjectSchema(schema)) {
    return panic(`@stll/folio-agents tool ${name} has no object input schema`);
  }
  return { description: definition.description, schema };
};

type AgentToolOptions = {
  agentTool: FolioAgentToolName;
  commands: readonly FolioCommandSpec[];
  extraProperties?: Readonly<Record<string, unknown>>;
};

const agentRead = ({
  agentTool,
  commands,
  extraProperties = {},
}: AgentToolOptions): FolioFileToolSpec => {
  const { description, schema } = agentDefinition(agentTool);
  return {
    type: "agentRead",
    name: agentTool,
    agentTool,
    description,
    argsSchema: {
      type: "object",
      properties: { ...schema.properties, ...extraProperties },
      required: schema.required,
      additionalProperties: false,
    },
    commands,
  };
};

/** Largest page `read_document` returns on any surface. */
export const MAX_READ_BLOCKS = 1000;

/**
 * Paging arguments the file surface adds to `read_document`: the agents tool
 * returns the whole body, which a large contract makes too big for one reply.
 */
const READ_DOCUMENT_PAGING_PROPERTIES = {
  maxBlocks: {
    type: "integer",
    minimum: 1,
    maximum: MAX_READ_BLOCKS,
    description: "Blocks per page. The response carries `nextCursor` when more remain.",
  },
  cursor: {
    type: "string",
    description:
      "`nextCursor` from the previous page. A cursor is bound to the fileVersion it was issued for.",
  },
} as const;

export const FOLIO_FILE_TOOLS: readonly FolioFileToolSpec[] = [
  agentRead({
    agentTool: FOLIO_AGENT_TOOL_NAMES.readDocument,
    commands: [{ name: "read", summary: "Print the document's blocks with ids and text hashes" }],
    extraProperties: READ_DOCUMENT_PAGING_PROPERTIES,
  }),
  agentRead({
    agentTool: FOLIO_AGENT_TOOL_NAMES.getDocumentOutline,
    commands: [{ name: "outline", summary: "Print the heading outline and section handles" }],
  }),
  agentRead({
    agentTool: FOLIO_AGENT_TOOL_NAMES.readSection,
    commands: [{ name: "section", summary: "Print one section's blocks by its outline handle" }],
  }),
  agentRead({
    agentTool: FOLIO_AGENT_TOOL_NAMES.listStories,
    commands: [{ name: "stories", summary: "List header, footer, and note stories" }],
  }),
  agentRead({
    agentTool: FOLIO_AGENT_TOOL_NAMES.readStory,
    commands: [{ name: "story", summary: "Print one story by its handle" }],
  }),
  agentRead({
    agentTool: FOLIO_AGENT_TOOL_NAMES.findText,
    commands: [{ name: "find", summary: "Find exact text and print range handles" }],
  }),
  agentRead({
    agentTool: FOLIO_AGENT_TOOL_NAMES.readComments,
    commands: [{ name: "comments", summary: "Print comment threads" }],
  }),
  agentRead({
    agentTool: FOLIO_AGENT_TOOL_NAMES.readChanges,
    commands: [{ name: "changes", summary: "Print pending tracked changes" }],
  }),
];

const toolsByName = new Map(FOLIO_FILE_TOOLS.map((tool) => [tool.name, tool]));

export const findFileTool = (name: string): FolioFileToolSpec | undefined => toolsByName.get(name);

export type FolioResolvedCommand = { tool: FolioFileToolSpec; command: FolioCommandSpec };

const commandsByName = new Map<string, FolioResolvedCommand>();
for (const tool of FOLIO_FILE_TOOLS) {
  for (const command of tool.commands) {
    if (commandsByName.has(command.name)) {
      panic(`two tools declare the command ${command.name}`);
    }
    commandsByName.set(command.name, { tool, command });
  }
}

export const findCommand = (name: string): FolioResolvedCommand | undefined =>
  commandsByName.get(name);

export const listCommands = (): readonly FolioResolvedCommand[] => [...commandsByName.values()];
