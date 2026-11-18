/**
 * The one tool vocabulary behind both surfaces. Every entry is an MCP tool
 * (its `name`, description, and envelope-wrapped input schema) and one or
 * more `folio` commands whose flags are generated from the same schema. Tools
 * that exist in `@stll/folio-agents` reuse its definitions verbatim; the file
 * envelope (`path`, `fileVersion`, ...) is added by each surface.
 */

import { panic } from "better-result";

import { getFolioToolDefinitions } from "@stll/folio-agents/tools";
import {
  FOLIO_AGENT_TOOL_NAMES,
  type FolioAgentToolName,
} from "@stll/folio-agents/types";

import { withBatchOperations } from "./batch-operations";
import type { DocumentPassName } from "./document-pass";
import { LABEL_SOURCES, NUMBERING_ACTIONS } from "./numbering";
import type { PackageRewriteName } from "./package-rewrite";
import { DEFAULT_PLACEHOLDER_PATTERNS } from "./placeholders";
import { DEFAULT_SANITIZE_TARGETS, SANITIZE_TARGETS } from "./sanitize";

export type JsonObjectSchema = {
  readonly type: "object";
  readonly properties: Readonly<Record<string, unknown>>;
  readonly required: readonly string[];
  readonly additionalProperties: false;
};

/** A positional argument after `<file>` that sets one tool argument. */
export type FolioPositionalSpec = {
  readonly name: string;
  readonly property: string;
};

/** One `folio <name>` command. */
export type FolioCommandSpec = {
  readonly name: string;
  readonly summary: string;
  /** Arguments the command implies (`accept` sets `action: "accept"`). */
  readonly preset?: Readonly<Record<string, unknown>>;
  /** Positionals after `<file>`, in order. */
  readonly positionals?: readonly FolioPositionalSpec[];
  /** Flag spellings that differ from the kebab-cased property (`ids` as `--id`). */
  readonly flagNames?: Readonly<Record<string, string>>;
  /** A bare JSON array passed to `--input` fills this property. */
  readonly inputArrayProperty?: string;
  /**
   * The command exits with `EXIT_CODES.unresolved` when its result counts
   * `unresolved` items, so a script can gate on it.
   */
  readonly exitWhenUnresolved?: true;
};

type FolioToolSpecBase = {
  readonly name: string;
  readonly description: string;
  /** The tool's own arguments, without the file envelope. */
  readonly argsSchema: JsonObjectSchema;
  readonly commands: readonly FolioCommandSpec[];
};

/**
 * How a tool runs. `agentRead` and `agentWrite` run a `@stll/folio-agents`
 * tool against one file (a write commits a transaction); `resolveChanges`
 * accepts or rejects tracked changes; `compare` diffs two files and, given a
 * destination, writes a redline; `packageRewrite` rewrites package parts
 * directly and, without a destination, reports what it would change;
 * `documentPass` inspects the opened document and, given a destination,
 * changes it through the reviewer.
 */
export type FolioFileToolSpec =
  | (FolioToolSpecBase & {
      readonly type: "agentRead";
      readonly agentTool: FolioAgentToolName;
    })
  | (FolioToolSpecBase & {
      readonly type: "agentWrite";
      readonly agentTool: FolioAgentToolName;
      /** Whether the caller may choose direct edits over tracked changes. */
      readonly editMode: "tracked-or-direct" | "fixed";
    })
  | (FolioToolSpecBase & { readonly type: "resolveChanges" })
  | (FolioToolSpecBase & { readonly type: "compare" })
  | (FolioToolSpecBase & {
      readonly type: "packageRewrite";
      readonly rewrite: PackageRewriteName;
    })
  | (FolioToolSpecBase & {
      readonly type: "documentPass";
      readonly pass: DocumentPassName;
      /** Whether the caller may choose direct edits over tracked changes. */
      readonly editMode: "tracked-or-direct" | "fixed";
    });

/** Whether a tool call may write, which decides its envelope and flags. */
export type FolioToolAccess = "read" | "write" | "readOrWrite";

export const toolAccess = (tool: FolioFileToolSpec): FolioToolAccess => {
  switch (tool.type) {
    case "agentRead":
      return "read";
    case "agentWrite":
    case "resolveChanges":
      return "write";
    case "compare":
    case "packageRewrite":
    case "documentPass":
      return "readOrWrite";
    default: {
      const unreachable: never = tool;
      return panic("Unhandled tool type", { unreachable });
    }
  }
};

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
    agentDefinitions.get(name) ??
    panic(`@stll/folio-agents has no tool named ${name}`);
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

const agentArgsSchema = (
  schema: JsonObjectSchema,
  extraProperties: Readonly<Record<string, unknown>>,
): JsonObjectSchema => ({
  type: "object",
  properties: { ...schema.properties, ...extraProperties },
  required: schema.required,
  additionalProperties: false,
});

const agentRead = ({
  agentTool,
  commands,
  extraProperties = {},
}: AgentToolOptions) => {
  const { description, schema } = agentDefinition(agentTool);
  return {
    type: "agentRead",
    name: agentTool,
    agentTool,
    description,
    argsSchema: agentArgsSchema(schema, extraProperties),
    commands,
  } as const satisfies FolioFileToolSpec;
};

const agentWrite = ({
  agentTool,
  commands,
  editMode,
}: AgentToolOptions & { editMode: "tracked-or-direct" | "fixed" }) => {
  const { description, schema } = agentDefinition(agentTool);
  const argsSchema = agentArgsSchema(schema, {});
  return {
    type: "agentWrite",
    name: agentTool,
    agentTool,
    editMode,
    description,
    argsSchema:
      agentTool === FOLIO_AGENT_TOOL_NAMES.suggestChanges
        ? withBatchOperations(argsSchema)
        : argsSchema,
    commands,
  } as const satisfies FolioFileToolSpec;
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
    description:
      "Blocks per page. The response carries `nextCursor` when more remain.",
  },
  cursor: {
    type: "string",
    description:
      "`nextCursor` from the previous page. A cursor is bound to the fileVersion it was issued for.",
  },
} as const;

export const RESOLVE_CHANGE_ACTIONS = ["accept", "reject"] as const;

export type ResolveChangeAction = (typeof RESOLVE_CHANGE_ACTIONS)[number];

const RESOLVE_CHANGES_TOOL = {
  type: "resolveChanges",
  name: "resolve_changes",
  description:
    "Accept or reject pending tracked changes. Pass the ids from `read_changes`, or `all: true` for every " +
    "change in every story; exactly one of the two. Accepting keeps an insertion and drops a deletion; " +
    "rejecting does the reverse.",
  argsSchema: {
    type: "object",
    properties: {
      action: {
        type: "string",
        enum: RESOLVE_CHANGE_ACTIONS,
        description: "accept or reject.",
      },
      ids: {
        type: "array",
        items: { type: "string" },
        minItems: 1,
        description: "Change ids from read_changes.",
      },
      all: { type: "boolean", description: "Resolve every pending change." },
    },
    required: ["action"],
    additionalProperties: false,
  },
  commands: [
    {
      name: "accept",
      summary: "Accept tracked changes by id, or all of them",
      preset: { action: "accept" },
      flagNames: { ids: "id" },
    },
    {
      name: "reject",
      summary: "Reject tracked changes by id, or all of them",
      preset: { action: "reject" },
      flagNames: { ids: "id" },
    },
  ],
} as const satisfies FolioFileToolSpec;

const COMPARE_TOOL = {
  type: "compare",
  name: "compare_documents",
  description:
    "Compare a base `.docx` (`path`) with a revised one (`revisedPath`). Without a destination, returns the " +
    "block-level differences. With one, writes the base package carrying the differences as tracked changes " +
    "(a redline) there; the inputs are never modified.",
  argsSchema: {
    type: "object",
    properties: {
      revisedPath: { type: "string", description: "The revised .docx." },
      revisedFileVersion: {
        type: "string",
        description:
          "Refuse unless the revised file still has this SHA-256 fileVersion.",
      },
    },
    required: ["revisedPath"],
    additionalProperties: false,
  },
  commands: [
    {
      name: "compare",
      summary: "Diff two files, or write their redline with -o",
      positionals: [{ name: "revised", property: "revisedPath" }],
    },
  ],
} as const satisfies FolioFileToolSpec;

const SANITIZE_TOOL = {
  type: "packageRewrite",
  name: "sanitize_document",
  rewrite: "sanitize",
  description:
    "Remove what a document carries about its past before it is reused. `comments`: every comment with its " +
    "parts (comments, commentsExtended, commentsIds, commentsExtensible), its anchors and references in every " +
    "story, and the people.xml entries no remaining revision names. `rsids`: every `w:rsid*` attribute, style " +
    "`w:rsid`, and the settings `w:rsids` table. `properties`: the creator and lastModifiedBy core properties " +
    "and the Company and Manager extended properties. Without a destination, reports what it would remove " +
    "and writes nothing.",
  argsSchema: {
    type: "object",
    properties: {
      remove: {
        type: "array",
        items: { type: "string", enum: SANITIZE_TARGETS },
        minItems: 1,
        description: `What to remove (default: ${DEFAULT_SANITIZE_TARGETS.join(", ")}).`,
      },
    },
    required: [],
    additionalProperties: false,
  },
  commands: [
    {
      name: "sanitize",
      summary: "Remove comments, rsids, and identifying properties",
    },
  ],
} as const satisfies FolioFileToolSpec;

const LANGUAGE_TAG_DESCRIPTION = "A BCP 47 language tag such as cs-CZ.";

const PROOFING_LANGUAGE_TOOL = {
  type: "packageRewrite",
  name: "proofing_language",
  rewrite: "proofingLanguage",
  description:
    "Report the proofing languages (`w:lang` val, eastAsia, bidi, and settings `w:themeFontLang`) each part " +
    "states, and with `set` write a language: into the `w:docDefaults` run properties, the default paragraph " +
    "style, every `w:lang` a style states, and `w:themeFontLang`. With `runs`, every `w:lang` a run or " +
    "paragraph mark states in the body, headers, footers, and notes is set too. Without a destination, " +
    "reports what it would change and writes nothing.",
  argsSchema: {
    type: "object",
    properties: {
      set: {
        type: "string",
        description: `The language for \`w:lang\` val. ${LANGUAGE_TAG_DESCRIPTION}`,
      },
      eastAsia: {
        type: "string",
        description: `The language for \`w:lang\` eastAsia. ${LANGUAGE_TAG_DESCRIPTION}`,
      },
      bidi: {
        type: "string",
        description: `The language for \`w:lang\` bidi. ${LANGUAGE_TAG_DESCRIPTION}`,
      },
      runs: {
        type: "boolean",
        description:
          "Also set every `w:lang` a run or paragraph mark states in the stories.",
      },
    },
    required: [],
    additionalProperties: false,
  },
  commands: [
    {
      name: "lang",
      summary: "Report or set the proofing language",
    },
  ],
} as const satisfies FolioFileToolSpec;

const PLACEHOLDERS_TOOL = {
  type: "documentPass",
  name: "placeholders",
  pass: "placeholders",
  editMode: "tracked-or-direct",
  description:
    "List the placeholders a draft still carries, in the body, headers, footers, and notes, with block ids and " +
    "context; with `fill` and a destination, replace them, as tracked changes unless direct edits are asked " +
    "for. A pattern's `name` group names the field that fills it (`{{party}}` takes `party`); a value keyed by " +
    "a placeholder's exact text (`[•]`) fills it too. The result counts what stays `unresolved`.",
  argsSchema: {
    type: "object",
    properties: {
      fill: {
        type: "object",
        additionalProperties: { type: "string" },
        description:
          "Values by field name or by exact placeholder text. Without a destination, shows what would fill each.",
      },
      patterns: {
        type: "array",
        items: { type: "string" },
        minItems: 1,
        description: `Regular expressions that replace the defaults (${DEFAULT_PLACEHOLDER_PATTERNS.join(" ")}).`,
      },
    },
    required: [],
    additionalProperties: false,
  },
  commands: [
    {
      name: "placeholders",
      summary: "List unresolved placeholders, or fill them",
      flagNames: { patterns: "pattern" },
      exitWhenUnresolved: true,
    },
  ],
} as const satisfies FolioFileToolSpec;

const NUMBERING_TOOL = {
  type: "packageRewrite",
  name: "numbering",
  rewrite: "numbering",
  description:
    "`literal`: write each numbered paragraph's label as text at its start and remove its list membership " +
    "(`w:numPr`, or `numId` 0 over a style's numbering), keeping the level's indentation, tab stops, label " +
    "run properties and `w:suff` separator as direct formatting. `labels: displayed` writes the label the " +
    "document shows; `continuous` counts each level 1, 2, 3 across the whole body, ignoring restarts. " +
    "`format` shapes the label from {label}, {n} (the level's count) and {path} (1.2.3). Bullets are left " +
    "alone; `numIds` and `blockIds` narrow the paragraphs. `apply`: put `blockIds` into numbering instance " +
    "`numId` at `level`, removing a leading label `strip` matches. Without a destination, reports the " +
    "paragraphs and labels and writes nothing.",
  argsSchema: {
    type: "object",
    properties: {
      action: {
        type: "string",
        enum: NUMBERING_ACTIONS,
        description: "literal (default) or apply.",
      },
      labels: {
        type: "string",
        enum: LABEL_SOURCES,
        description: "literal: displayed (default) or continuous.",
      },
      format: {
        type: "string",
        description:
          'literal: the label template, e.g. "[{n}.]" (default {label}, or {n}. for continuous).',
      },
      numIds: {
        type: "array",
        items: { type: "integer" },
        description: "literal: only paragraphs in these numbering instances.",
      },
      blockIds: {
        type: "array",
        items: { type: "string" },
        description:
          "Paragraphs to change, by block id (literal: default all numbered).",
      },
      numId: {
        type: "integer",
        description: "apply: the numbering instance.",
      },
      level: {
        type: "integer",
        minimum: 0,
        maximum: 8,
        description: "apply: the zero-based level (default 0).",
      },
      strip: {
        type: "string",
        description:
          "apply: a regular expression for a leading literal label to remove, e.g. \\[\\d+\\.\\]\\t.",
      },
    },
    required: [],
    additionalProperties: false,
  },
  commands: [
    {
      name: "literal-numbering",
      summary: "Write list numbers as literal text",
      preset: { action: "literal" },
      flagNames: { blockIds: "block-id" },
    },
    {
      name: "apply-numbering",
      summary: "Put paragraphs into a numbering instance",
      preset: { action: "apply" },
      flagNames: { blockIds: "block-id" },
    },
  ],
} as const satisfies FolioFileToolSpec;

export const FOLIO_FILE_TOOLS: readonly FolioFileToolSpec[] = [
  agentRead({
    agentTool: FOLIO_AGENT_TOOL_NAMES.readDocument,
    commands: [
      {
        name: "read",
        summary: "Print the document's blocks with ids and text hashes",
      },
    ],
    extraProperties: READ_DOCUMENT_PAGING_PROPERTIES,
  }),
  agentRead({
    agentTool: FOLIO_AGENT_TOOL_NAMES.getDocumentOutline,
    commands: [
      {
        name: "outline",
        summary: "Print the heading outline and section handles",
      },
    ],
  }),
  agentRead({
    agentTool: FOLIO_AGENT_TOOL_NAMES.readSection,
    commands: [
      {
        name: "section",
        summary: "Print one section's blocks by its outline handle",
      },
    ],
  }),
  agentRead({
    agentTool: FOLIO_AGENT_TOOL_NAMES.listStories,
    commands: [
      { name: "stories", summary: "List header, footer, and note stories" },
    ],
  }),
  agentRead({
    agentTool: FOLIO_AGENT_TOOL_NAMES.readStory,
    commands: [{ name: "story", summary: "Print one story by its handle" }],
  }),
  agentRead({
    agentTool: FOLIO_AGENT_TOOL_NAMES.findText,
    commands: [
      { name: "find", summary: "Find exact text and print range handles" },
    ],
  }),
  agentRead({
    agentTool: FOLIO_AGENT_TOOL_NAMES.readComments,
    commands: [{ name: "comments", summary: "Print comment threads" }],
  }),
  agentRead({
    agentTool: FOLIO_AGENT_TOOL_NAMES.readChanges,
    commands: [{ name: "changes", summary: "Print pending tracked changes" }],
  }),
  agentWrite({
    agentTool: FOLIO_AGENT_TOOL_NAMES.suggestChanges,
    editMode: "tracked-or-direct",
    commands: [
      {
        name: "suggest",
        summary: "Apply a batch of edit operations as tracked changes",
        inputArrayProperty: "operations",
      },
    ],
  }),
  agentWrite({
    agentTool: FOLIO_AGENT_TOOL_NAMES.addComment,
    editMode: "fixed",
    commands: [
      {
        name: "comment",
        summary: "Comment on a block, optionally quoting text in it",
      },
    ],
  }),
  agentWrite({
    agentTool: FOLIO_AGENT_TOOL_NAMES.replyComment,
    editMode: "fixed",
    commands: [{ name: "reply", summary: "Reply to a comment thread" }],
  }),
  agentWrite({
    agentTool: FOLIO_AGENT_TOOL_NAMES.resolveComment,
    editMode: "fixed",
    commands: [
      { name: "resolve", summary: "Resolve or reopen a comment thread" },
    ],
  }),
  RESOLVE_CHANGES_TOOL,
  COMPARE_TOOL,
  SANITIZE_TOOL,
  PROOFING_LANGUAGE_TOOL,
  PLACEHOLDERS_TOOL,
  NUMBERING_TOOL,
];

const toolsByName = new Map(FOLIO_FILE_TOOLS.map((tool) => [tool.name, tool]));

export const findFileTool = (name: string): FolioFileToolSpec | undefined =>
  toolsByName.get(name);

export type FolioResolvedCommand = {
  tool: FolioFileToolSpec;
  command: FolioCommandSpec;
};

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

export const listCommands = (): readonly FolioResolvedCommand[] => [
  ...commandsByName.values(),
];
