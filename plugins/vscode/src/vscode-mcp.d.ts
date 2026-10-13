import "vscode";

// MCP surface used by Folio, copied from VS Code 1.101.0's vscode.d.ts:
// https://github.com/microsoft/vscode/blob/1.101.0/src/vscode-dts/vscode.d.ts
// Keep the base typings at the editor floor. Registration is optional at runtime.
declare module "vscode" {
  export class McpStdioServerDefinition {
    readonly label: string;
    cwd?: Uri;
    command: string;
    args: string[];
    env: Record<string, string | number | null>;
    version?: string;
    constructor(
      label: string,
      command: string,
      args?: string[],
      env?: Record<string, string | number | null>,
      version?: string,
    );
  }

  export type McpServerDefinitionProvider<
    T extends McpStdioServerDefinition = McpStdioServerDefinition,
  > = {
    readonly onDidChangeMcpServerDefinitions?: Event<void>;
    provideMcpServerDefinitions(token: CancellationToken): ProviderResult<T[]>;
  }

  export namespace lm {
    export const registerMcpServerDefinitionProvider:
      | ((id: string, provider: McpServerDefinitionProvider) => Disposable)
      | undefined;
  }
}
