import { $ } from "bun";
import { readFileSync } from "node:fs";
import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

/** Copy only the registry release policy; workspace linker settings stay in the repo. */
export const writeCleanRoomInstallConfig = async (
  consumerDir: string,
  linker?: "isolated",
): Promise<void> => {
  const { panic } = await import("better-result");
  const config: unknown = Bun.TOML.parse(
    await readFile(new URL("../bunfig.toml", import.meta.url), "utf8"),
  );
  if (!isRecord(config) || !isRecord(config["install"])) {
    return panic("bunfig.toml must declare an install policy");
  }
  const { minimumReleaseAge, minimumReleaseAgeExcludes } = config["install"];
  if (
    typeof minimumReleaseAge !== "number" ||
    !Number.isSafeInteger(minimumReleaseAge) ||
    minimumReleaseAge <= 0 ||
    !Array.isArray(minimumReleaseAgeExcludes) ||
    !minimumReleaseAgeExcludes.every(
      (name: unknown) => typeof name === "string",
    )
  ) {
    return panic(
      "bunfig.toml must declare a positive release age and package exclusions",
    );
  }
  await writeFile(
    path.join(consumerDir, "bunfig.toml"),
    `[install]\n${linker ? `linker = "${linker}"\n` : ""}minimumReleaseAge = ${minimumReleaseAge}\nminimumReleaseAgeExcludes = ${JSON.stringify(minimumReleaseAgeExcludes)}\n`,
  );
};

/** Bootstrap a locked repository environment using its complete install policy. */
export const installFrozenRepositoryDependencies = (repoRoot: string): void => {
  const config: unknown = Bun.TOML.parse(
    readFileSync(path.join(repoRoot, "bunfig.toml"), "utf8"),
  );
  if (
    !isRecord(config) ||
    !isRecord(config["install"]) ||
    typeof config["install"]["minimumReleaseAge"] !== "number" ||
    !Number.isSafeInteger(config["install"]["minimumReleaseAge"]) ||
    config["install"]["minimumReleaseAge"] <= 0
  )
    throw new TypeError(
      "Frozen repository installs require the release-age policy",
    );
  const result = Bun.spawnSync(
    [process.execPath, "install", "--frozen-lockfile"],
    {
      cwd: repoRoot,
      stdout: "inherit",
      stderr: "inherit",
    },
  );
  if (result.exitCode !== 0)
    throw new TypeError("Frozen repository install failed");
};

type CleanRoomInstallOptions = {
  consumerDir: string;
  dependencies: readonly string[];
  linker?: "isolated";
};

/** Own the install so every invocation writes the current repository policy first. */
export const installCleanRoomDependencies = async ({
  consumerDir,
  dependencies,
  linker,
}: CleanRoomInstallOptions): Promise<void> => {
  await writeCleanRoomInstallConfig(consumerDir, linker);
  await $`bun add ${dependencies}`.cwd(consumerDir).quiet();
};
