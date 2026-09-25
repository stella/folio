#!/usr/bin/env bun
// Keep the plugins on the @stll/folio-cli version they run: each plugin's
// manifest `version` and every `@stll/folio-cli@<version>` spec in its files
// carry the CLI package's version. `--write` updates them (run by
// `changeset:version`); without it the script reports drift.

import { panic } from "better-result";
import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";

const repoRoot = path.resolve(import.meta.dir, "..");
export const CLI_PACKAGE_JSON = path.join(repoRoot, "packages", "cli", "package.json");

const CLI_SPEC = /@stll\/folio-cli@[0-9A-Za-z.+-]+/gu;

/** A file that names the CLI version, and how its own version field is spelled. */
type VersionedFile = {
  readonly file: string;
  /** Groups: the text before the manifest's version, the version, the text after. */
  readonly manifestVersion?: RegExp;
};

export const PLUGIN_VERSION_FILES: readonly VersionedFile[] = [
  {
    file: path.join(repoRoot, "plugins", "claude-code", ".claude-plugin", "plugin.json"),
    manifestVersion: /("version":\s*")([^"]*)(")/u,
  },
  { file: path.join(repoRoot, "plugins", "claude-code", ".mcp.json") },
  {
    file: path.join(repoRoot, "plugins", "herdr", "herdr-plugin.toml"),
    manifestVersion: /^(version = ")([^"]*)(")/mu,
  },
  { file: path.join(repoRoot, "plugins", "herdr", "bin", "folio-command.sh") },
];

export const cliVersion = async (): Promise<string> => {
  const parsed: unknown = JSON.parse(await readFile(CLI_PACKAGE_JSON, "utf8"));
  const version =
    typeof parsed === "object" && parsed !== null && "version" in parsed
      ? parsed.version
      : undefined;
  return typeof version === "string" ? version : panic("@stll/folio-cli has no string version");
};

/** Every version a plugin file names: its manifest version and each CLI spec. */
export const pluginVersions = async (): Promise<{ file: string; versions: string[] }[]> => {
  const found: { file: string; versions: string[] }[] = [];
  for (const { file, manifestVersion } of PLUGIN_VERSION_FILES) {
    const text = await readFile(file, "utf8");
    const versions = (text.match(CLI_SPEC) ?? []).map((spec) =>
      spec.slice("@stll/folio-cli@".length),
    );
    if (manifestVersion !== undefined) {
      versions.push(manifestVersion.exec(text)?.[2] ?? panic(`${file} has no version field`));
    }
    if (versions.length === 0) panic(`${file} names no @stll/folio-cli version`);
    found.push({ file: path.relative(repoRoot, file), versions });
  }
  return found;
};

const writeVersions = async (version: string): Promise<void> => {
  for (const { file, manifestVersion } of PLUGIN_VERSION_FILES) {
    let text = (await readFile(file, "utf8")).replace(CLI_SPEC, `@stll/folio-cli@${version}`);
    if (manifestVersion !== undefined) text = text.replace(manifestVersion, `$1${version}$3`);
    await writeFile(file, text);
  }
};

if (import.meta.main) {
  const write = process.argv[2] === "--write";
  if (process.argv.length > 3 || (process.argv[2] !== undefined && !write)) {
    panic("usage: bun scripts/sync-plugin-cli-version.ts [--write]");
  }
  const version = await cliVersion();
  if (write) {
    await writeVersions(version);
    console.log(`Synchronized the plugins to @stll/folio-cli ${version}`);
  } else {
    const drift = (await pluginVersions()).filter(({ versions }) =>
      versions.some((found) => found !== version),
    );
    if (drift.length > 0) {
      panic(
        `Plugin version drift from @stll/folio-cli ${version}: ${drift.map(({ file, versions }) => `${file}=${versions.join(",")}`).join("; ")}`,
      );
    }
    console.log(`The plugins are on @stll/folio-cli ${version}`);
  }
}
