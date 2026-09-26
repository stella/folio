/**
 * How the extension launches the folio MCP server: the CLI bundled in the
 * extension, run by the editor's own Node.js, with the workspace folders as
 * the allowed roots. Nothing here imports `vscode`, so it is unit-tested.
 */

import { cliCommand, type CliRuntime } from "./runtime";

/** The id under `contributes.mcpServerDefinitionProviders`. */
export const FOLIO_MCP_PROVIDER_ID = "folio.mcp";

export const FOLIO_MCP_LABEL = "Folio";

/** What `McpStdioServerDefinition` is built from. */
export type McpLaunch = {
  readonly label: string;
  readonly command: string;
  readonly args: readonly string[];
  readonly env: Readonly<Record<string, string>>;
  /** The first root: the base a relative tool path resolves against. */
  readonly cwd: string;
  readonly version: string;
};

export type McpLaunchOptions = {
  readonly runtime: CliRuntime;
  /** Absolute paths of the workspace folders on disk, in workspace order. */
  readonly roots: readonly string[];
  /** The `folio.author` setting. Blank leaves the server to its own fallbacks. */
  readonly author: string | undefined;
  /** Changing it tells the editor the server's tools may have changed. */
  readonly version: string;
};

/** The author to pass, or `undefined` when the setting is blank. */
export const configuredAuthor = (author: string | undefined): string | undefined => {
  const trimmed = author?.trim() ?? "";
  return trimmed === "" ? undefined : trimmed;
};

/**
 * The server launch, or `null` without a folder to root it in: the server
 * refuses every path outside its roots, so it has nothing to work on.
 */
export const buildMcpLaunch = ({
  runtime,
  roots,
  author,
  version,
}: McpLaunchOptions): McpLaunch | null => {
  const unique = [...new Set(roots)];
  const first = unique.at(0);
  if (first === undefined) return null;
  const name = configuredAuthor(author);
  const { command, args, env } = cliCommand(runtime, [
    "mcp",
    ...unique.flatMap((root) => ["--root", root]),
    ...(name === undefined ? [] : ["--author", name]),
  ]);
  return { label: FOLIO_MCP_LABEL, command, args, env, cwd: first, version };
};
