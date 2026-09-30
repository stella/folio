// Activation against a stand-in `vscode` module: what the extension registers,
// and the MCP server definition it hands the editor.

import { afterEach, beforeEach, describe, expect, mock, test } from "bun:test";

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
  provideMcpServerDefinitions: () => Promise<McpStdioServerDefinition[]>;
  resolveMcpServerDefinition?: unknown;
};

const state = {
  trusted: true,
  folders: [] as { uri: { scheme: string; fsPath: string } }[],
  author: "" as string | undefined,
  git: undefined as string | undefined,
  os: undefined as string | undefined,
  warnings: 0,
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
    showWarningMessage: () => {
      state.warnings += 1;
      return Promise.resolve(undefined);
    },
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

const actualSettings = await import("./editor-settings");
void mock.module("./editor-settings", () => ({
  ...actualSettings,
  readGitUserName: () => Promise.resolve(state.git),
  osUserName: () => state.os,
}));

const { activate } = await import("./extension");

const context = {
  subscriptions: [] as unknown[],
  extensionUri: { path: "/ext" },
  extension: { packageJSON: { version: "9.8.7" } },
  asAbsolutePath: (relative: string) => `/ext/${relative}`,
};

const inheritedAuthor = process.env["FOLIO_AUTHOR"];

beforeEach(() => {
  delete process.env["FOLIO_AUTHOR"];
  state.trusted = true;
  state.folders = [{ uri: { scheme: "file", fsPath: "/work/contracts" } }];
  state.author = "Ada Lovelace";
  state.git = undefined;
  state.os = undefined;
  state.warnings = 0;
  state.providers.clear();
  state.editors.clear();
  state.commands.clear();
  state.onFolders = [];
  activate(context as never);
});

afterEach(() => {
  if (inheritedAuthor === undefined) delete process.env["FOLIO_AUTHOR"];
  else process.env["FOLIO_AUTHOR"] = inheritedAuthor;
});

const definitions = async () =>
  (await state.providers.get("folio.mcp")?.provideMcpServerDefinitions()) ?? [];

describe("activate", () => {
  test("registers the editor, its read-only command, and the MCP provider", () => {
    expect([...state.editors.keys()]).toEqual(["folio.docxEditor"]);
    expect([...state.commands]).toEqual(["folio.openReadOnly"]);
    expect([...state.providers.keys()]).toEqual(["folio.mcp"]);
  });

  test("exports test hooks only under FOLIO_VSCODE_TEST", () => {
    expect(activate(context as never)).toBeUndefined();
    process.env["FOLIO_VSCODE_TEST"] = "1";
    try {
      expect(activate(context as never)).toBeDefined();
    } finally {
      delete process.env["FOLIO_VSCODE_TEST"];
    }
  });

  test("keeps one editor per document, alive while hidden", () => {
    expect(state.editors.get("folio.docxEditor")?.options).toEqual({
      supportsMultipleEditorsPerDocument: false,
      webviewOptions: { retainContextWhenHidden: true },
    });
  });

  test("defines the server as the bundled CLI under the editor's Node.js", async () => {
    const [definition] = await definitions();

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

  test("offers no server in an untrusted workspace", async () => {
    state.trusted = false;

    expect(await definitions()).toEqual([]);
  });

  test("roots the server only in folders on disk", async () => {
    state.folders = [
      { uri: { scheme: "vscode-vfs", fsPath: "/remote" } },
      { uri: { scheme: "file", fsPath: "/work/b" } },
    ];

    expect(await (await definitions())[0]?.args).toEqual([
      "/ext/dist/cli/folio.mjs",
      "mcp",
      "--root",
      "/work/b",
      "--author",
      "Ada Lovelace",
    ]);
  });

  test("offers no server without a folder, and says so when folders change", async () => {
    state.folders = [];
    let changes = 0;
    state.providers.get("folio.mcp")?.onDidChangeMcpServerDefinitions(() => {
      changes += 1;
    });

    expect(await definitions()).toEqual([]);
    for (const listener of state.onFolders) listener();
    expect(changes).toBe(1);
  });

  test("names agent changes by the setting before git and the OS", async () => {
    state.git = "Git Name";
    state.os = "osuser";

    expect((await definitions())[0]?.args.slice(-2)).toEqual(["--author", "Ada Lovelace"]);
  });

  test("falls back to git user.name when the setting is blank", async () => {
    state.author = "  ";
    state.git = "Git Name";
    state.os = "osuser";

    expect((await definitions())[0]?.args.slice(-2)).toEqual(["--author", "Git Name"]);
  });

  test("falls back to the OS user name without a setting or git name", async () => {
    state.author = "";
    state.os = "osuser";

    expect((await definitions())[0]?.args.slice(-2)).toEqual(["--author", "osuser"]);
    expect(state.warnings).toBe(0);
  });

  test("never asks for folio.author", async () => {
    state.author = "";
    await definitions();
    const provider = state.providers.get("folio.mcp");

    expect(provider?.resolveMcpServerDefinition).toBeUndefined();
    expect(state.warnings).toBe(0);
  });
});
