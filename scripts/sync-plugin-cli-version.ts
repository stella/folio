#!/usr/bin/env bun
// Keep the Claude Code plugin on the @stll/folio-cli version it runs: the
// plugin manifest's `version` and the package spec in its MCP server command
// both carry the CLI package's version. `--write` updates them (run by
// `changeset:version`); without it the script reports drift.

import { panic } from "better-result";
import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";

const repoRoot = path.resolve(import.meta.dir, "..");
export const CLI_PACKAGE_JSON = path.join(repoRoot, "packages", "cli", "package.json");
export const PLUGIN_MANIFEST = path.join(
  repoRoot,
  "plugins",
  "claude-code",
  ".claude-plugin",
  "plugin.json",
);
export const PLUGIN_MCP_CONFIG = path.join(repoRoot, "plugins", "claude-code", ".mcp.json");

const CLI_SPEC = /@stll\/folio-cli@[^"\s]+/gu;

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const readJson = async (file: string): Promise<Record<string, unknown>> => {
  const parsed: unknown = JSON.parse(await readFile(file, "utf8"));
  return isRecord(parsed) ? parsed : panic(`${file} is not a JSON object`);
};

export const cliVersion = async (): Promise<string> => {
  const { version } = await readJson(CLI_PACKAGE_JSON);
  return typeof version === "string" ? version : panic("@stll/folio-cli has no string version");
};

/** The versions the plugin files name, for comparison with the CLI's. */
export const pluginVersions = async (): Promise<{ manifest: unknown; mcp: string[] }> => {
  const { version } = await readJson(PLUGIN_MANIFEST);
  const mcp = (await readFile(PLUGIN_MCP_CONFIG, "utf8")).match(CLI_SPEC) ?? [];
  return { manifest: version, mcp: mcp.map((spec) => spec.slice("@stll/folio-cli@".length)) };
};

const writeVersions = async (version: string): Promise<void> => {
  const manifest = await readFile(PLUGIN_MANIFEST, "utf8");
  await writeFile(PLUGIN_MANIFEST, manifest.replace(/("version":\s*")[^"]*(")/u, `$1${version}$2`));
  const mcp = await readFile(PLUGIN_MCP_CONFIG, "utf8");
  await writeFile(PLUGIN_MCP_CONFIG, mcp.replace(CLI_SPEC, `@stll/folio-cli@${version}`));
};

if (import.meta.main) {
  const write = process.argv[2] === "--write";
  if (process.argv.length > 3 || (process.argv[2] !== undefined && !write)) {
    panic("usage: bun scripts/sync-plugin-cli-version.ts [--write]");
  }
  const version = await cliVersion();
  if (write) {
    await writeVersions(version);
    console.log(`Synchronized the Claude Code plugin to @stll/folio-cli ${version}`);
  } else {
    const found = await pluginVersions();
    if (
      found.manifest !== version ||
      found.mcp.length === 0 ||
      found.mcp.some((v) => v !== version)
    ) {
      panic(
        `Plugin version drift: cli=${version}, plugin.json=${String(found.manifest)}, .mcp.json=${found.mcp.join(",")}`,
      );
    }
    console.log(`Claude Code plugin is on @stll/folio-cli ${version}`);
  }
}
