// Activation against a stand-in `vscode` module: what the extension registers,
// and the MCP server definition it hands the editor.

import { beforeEach, describe, expect, mock, test } from "bun:test";

type Listener = () => void;

class EventEmitter {
  listeners: Listener[] = [];
  event = (listener: Listener) => {
    this.listeners.push(listener);
    return { dispose: () => undefined };
  };
  fire = () => {
    for (const listener of this.listeners) listener();
  };
  dispose = () => undefined;
}

class McpStdioServerDefinition {
  cwd: { fsPath: string } | undefined;
  readonly label: string;
  readonly command: string;
  readonly args: string[];
  readonly env: Record<string, string>;
  readonly version: string;
  constructor(
    label: string,
    command: string,
    args: string[],
    env: Record<string, string>,
    version: string,
  ) {
    this.label = label;
    this.command = command;
    this.args = args;
    this.env = env;
    this.version = version;
  }
}

type Provider = {
  onDidChangeMcpServerDefinitions: (listener: Listener) => unknown;
  provideMcpServerDefinitions: () => McpStdioServerDefinition[];
};

const state = {
  trusted: true,
  folders: [] as { uri: { scheme: string; fsPath: string } }[],
  author: "" as string | undefined,
  providers: new Map<string, Provider>(),
  editors: new Map<string, { options: unknown }>(),
  commands: new Set<string>(),
  onFolders: [] as Listener[],
};

const disposable = { dispose: () => undefined };

void mock.module("vscode", () => ({
  EventEmitter,
  McpStdioServerDefinition,
  Disposable: { from: () => disposable },
  Uri: {
    file: (fsPath: string) => ({ fsPath }),
    joinPath: (base: { path: string }, ...parts: string[]) => ({
      path: [base.path, ...parts].join("/"),
    }),
  },
  lm: {
    registerMcpServerDefinitionProvider: (id: string, provider: Provider) => {
      state.providers.set(id, provider);
      return disposable;
    },
  },
  window: {
    registerCustomEditorProvider: (viewType: string, _provider: unknown, options: unknown) => {
      state.editors.set(viewType, { options });
      return disposable;
    },
  },
  commands: {
    registerCommand: (id: string) => {
      state.commands.add(id);
      return disposable;
    },
  },
  workspace: {
    get isTrusted() {
      return state.trusted;
    },
    get workspaceFolders() {
      return state.folders;
    },
    getConfiguration: () => ({ get: () => state.author }),
    onDidChangeWorkspaceFolders: (listener: Listener) => {
      state.onFolders.push(listener);
      return disposable;
    },
    onDidGrantWorkspaceTrust: () => disposable,
    onDidChangeConfiguration: () => disposable,
  },
}));

const { activate } = await import("./extension");

const context = {
  subscriptions: [] as unknown[],
  extensionUri: { path: "/ext" },
  extension: { packageJSON: { version: "9.8.7" } },
  asAbsolutePath: (relative: string) => `/ext/${relative}`,
};

beforeEach(() => {
  state.trusted = true;
  state.folders = [{ uri: { scheme: "file", fsPath: "/work/contracts" } }];
  state.author = "Ada Lovelace";
  state.providers.clear();
  state.editors.clear();
  state.commands.clear();
  state.onFolders = [];
  activate(context as never);
});

const definitions = () => state.providers.get("folio.mcp")?.provideMcpServerDefinitions() ?? [];

describe("activate", () => {
  test("registers the editor, the preview, their commands, and the MCP provider", () => {
    expect([...state.editors.keys()]).toEqual(["folio.docxEditor", "folio.docxPreview"]);
    expect([...state.commands]).toEqual(["folio.openEditor", "folio.openPreview"]);
    expect([...state.providers.keys()]).toEqual(["folio.mcp"]);
  });

  test("keeps one editor per document, alive while hidden", () => {
    expect(state.editors.get("folio.docxEditor")?.options).toEqual({
      supportsMultipleEditorsPerDocument: false,
      webviewOptions: { retainContextWhenHidden: true },
    });
  });

  test("defines the server as the bundled CLI under the editor's Node.js", () => {
    const [definition] = definitions();

    expect(definition).toBeDefined();
    expect(definition?.label).toBe("Folio");
    expect(definition?.command).toBe(process.execPath);
    expect(definition?.args).toEqual([
      "/ext/dist/cli/folio.mjs",
      "mcp",
      "--root",
      "/work/contracts",
      "--author",
      "Ada Lovelace",
    ]);
    expect(definition?.env).toEqual({ ELECTRON_RUN_AS_NODE: "1" });
    expect(definition?.version).toBe("9.8.7");
    expect(definition?.cwd).toEqual({ fsPath: "/work/contracts" });
  });

  test("offers no server in an untrusted workspace", () => {
    state.trusted = false;

    expect(definitions()).toEqual([]);
  });

  test("roots the server only in folders on disk", () => {
    state.folders = [
      { uri: { scheme: "vscode-vfs", fsPath: "/remote" } },
      { uri: { scheme: "file", fsPath: "/work/b" } },
    ];

    expect(definitions()[0]?.args).toEqual([
      "/ext/dist/cli/folio.mjs",
      "mcp",
      "--root",
      "/work/b",
      "--author",
      "Ada Lovelace",
    ]);
  });

  test("offers no server without a folder, and says so when folders change", () => {
    state.folders = [];
    let changes = 0;
    state.providers.get("folio.mcp")?.onDidChangeMcpServerDefinitions(() => {
      changes += 1;
    });

    expect(definitions()).toEqual([]);
    for (const listener of state.onFolders) listener();
    expect(changes).toBe(1);
  });
});
