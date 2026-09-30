/**
 * Folio DOCX: the `.docx` editor, and the folio MCP server offered to the
 * editor's agent mode. Both run the folio CLI bundled in the extension.
 */

import path from "node:path";
import * as vscode from "vscode";

import { registerEditor, type EditorTestHooks } from "./editor";
import { osUserName, readGitUserName, resolveEditorAuthor } from "./editor-settings";
import { buildMcpLaunch, configuredAuthor, FOLIO_MCP_PROVIDER_ID } from "./mcp";
import type { CliRuntime } from "./runtime";

const AUTHOR_SETTING = "folio.author";

const extensionVersion = (context: vscode.ExtensionContext): string => {
  const packageJson: unknown = context.extension.packageJSON;
  if (typeof packageJson === "object" && packageJson !== null && "version" in packageJson) {
    const { version } = packageJson;
    if (typeof version === "string") return version;
  }
  return "0.0.0";
};

const authorSetting = (): string | undefined =>
  vscode.workspace.getConfiguration().get<string>(AUTHOR_SETTING);

/** Workspace folders on disk; the server cannot root itself in a virtual one. */
const diskRoots = (): string[] =>
  (vscode.workspace.workspaceFolders ?? [])
    .filter((folder) => folder.uri.scheme === "file")
    .map((folder) => folder.uri.fsPath);

/**
 * The name on agent changes, resolved as the editor does: the setting (or
 * `FOLIO_AUTHOR`), then git's `user.name` for the first root, then the OS
 * account name.
 */
const mcpAuthor = async (root: string): Promise<string> =>
  resolveEditorAuthor({
    setting: configuredAuthor(authorSetting()) ?? process.env["FOLIO_AUTHOR"],
    gitUserName: await readGitUserName(root),
    osUserName: osUserName(),
  });

const registerMcp = (context: vscode.ExtensionContext, runtime: CliRuntime): vscode.Disposable => {
  const changed = new vscode.EventEmitter<void>();
  const version = extensionVersion(context);

  const provider: vscode.McpServerDefinitionProvider<vscode.McpStdioServerDefinition> = {
    onDidChangeMcpServerDefinitions: changed.event,
    provideMcpServerDefinitions: async () => {
      // The server changes files; an untrusted workspace does not get it.
      if (!vscode.workspace.isTrusted) return [];
      const roots = diskRoots();
      const first = roots.at(0);
      if (first === undefined) return [];
      const launch = buildMcpLaunch({ runtime, roots, author: await mcpAuthor(first), version });
      if (launch === null) return [];
      const definition = new vscode.McpStdioServerDefinition(
        launch.label,
        launch.command,
        [...launch.args],
        { ...launch.env },
        launch.version,
      );
      definition.cwd = vscode.Uri.file(launch.cwd);
      return [definition];
    },
  };

  return vscode.Disposable.from(
    changed,
    vscode.lm.registerMcpServerDefinitionProvider(FOLIO_MCP_PROVIDER_ID, provider),
    vscode.workspace.onDidChangeWorkspaceFolders(() => changed.fire()),
    vscode.workspace.onDidGrantWorkspaceTrust(() => changed.fire()),
    vscode.workspace.onDidChangeConfiguration((event) => {
      if (event.affectsConfiguration(AUTHOR_SETTING)) changed.fire();
    }),
  );
};

/**
 * The extension's exports: nothing, except the editor's test hooks when the
 * smoke test runs it with `FOLIO_VSCODE_TEST=1`.
 */
export const activate = (context: vscode.ExtensionContext): EditorTestHooks | undefined => {
  const runtime: CliRuntime = {
    nodePath: process.execPath,
    cliEntry: context.asAbsolutePath(path.join("dist", "cli", "folio.mjs")),
  };
  const testMode = process.env["FOLIO_VSCODE_TEST"] === "1";
  const editor = registerEditor(context, runtime, testMode);
  context.subscriptions.push(editor.disposable, registerMcp(context, runtime));
  return testMode ? editor.testHooks : undefined;
};

export const deactivate = (): void => undefined;
