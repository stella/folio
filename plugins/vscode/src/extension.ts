/**
 * Folio DOCX: the `.docx` editor, a read-only preview, and the folio MCP
 * server offered to the editor's agent mode. All three run the folio CLI
 * bundled in the extension.
 */

import path from "node:path";
import * as vscode from "vscode";

import { registerEditor } from "./editor";
import { buildMcpLaunch, configuredAuthor, FOLIO_MCP_PROVIDER_ID } from "./mcp";
import { registerPreview } from "./preview";
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

/** Point the reader at `folio.author`; the server still starts without it. */
const suggestAuthorSetting = async (): Promise<void> => {
  const choice = await vscode.window.showWarningMessage(
    "Folio: set folio.author to the name recorded on tracked changes and comments. Until then the MCP server uses git user.name, and refuses changes if that is unset too.",
    "Open Settings",
  );
  if (choice === "Open Settings") {
    await vscode.commands.executeCommand("workbench.action.openSettings", AUTHOR_SETTING);
  }
};

const registerMcp = (context: vscode.ExtensionContext, runtime: CliRuntime): vscode.Disposable => {
  const changed = new vscode.EventEmitter<void>();
  const version = extensionVersion(context);
  let warnedAboutAuthor = false;

  const provider: vscode.McpServerDefinitionProvider<vscode.McpStdioServerDefinition> = {
    onDidChangeMcpServerDefinitions: changed.event,
    provideMcpServerDefinitions: () => {
      // The server changes files; an untrusted workspace does not get it.
      if (!vscode.workspace.isTrusted) return [];
      const launch = buildMcpLaunch({
        runtime,
        roots: diskRoots(),
        author: authorSetting(),
        version,
      });
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
    resolveMcpServerDefinition: (server) => {
      const unnamed =
        configuredAuthor(authorSetting()) === undefined &&
        configuredAuthor(process.env["FOLIO_AUTHOR"]) === undefined;
      if (unnamed && !warnedAboutAuthor) {
        warnedAboutAuthor = true;
        void suggestAuthorSetting();
      }
      return server;
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

export const activate = (context: vscode.ExtensionContext): void => {
  const runtime: CliRuntime = {
    nodePath: process.execPath,
    cliEntry: context.asAbsolutePath(path.join("dist", "cli", "folio.mjs")),
  };
  context.subscriptions.push(
    registerEditor(context, runtime),
    registerPreview(context, runtime),
    registerMcp(context, runtime),
  );
};

export const deactivate = (): void => undefined;
