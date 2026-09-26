/**
 * The folio CLI the extension runs: `dist/cli/folio.mjs`, bundled at build
 * time from packages/cli, executed by the editor's own Node.js so nothing is
 * downloaded and no Node.js install is needed.
 */

export type CliRuntime = {
  /** The editor's executable (`process.execPath` in the extension host). */
  readonly nodePath: string;
  /** Absolute path of the bundled CLI entry. */
  readonly cliEntry: string;
};

export type CliCommand = {
  readonly command: string;
  readonly args: readonly string[];
  readonly env: Readonly<Record<string, string>>;
};

/**
 * The command line for `folio <args>`. The editor's executable is an Electron
 * binary on the desktop; `ELECTRON_RUN_AS_NODE` makes it behave as Node.js. A
 * remote extension host already runs plain Node.js, which ignores the variable.
 */
export const cliCommand = (runtime: CliRuntime, args: readonly string[]): CliCommand => ({
  command: runtime.nodePath,
  args: [runtime.cliEntry, ...args],
  env: { ELECTRON_RUN_AS_NODE: "1" },
});
