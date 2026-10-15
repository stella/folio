import { expect, test } from "bun:test";
import { mkdir, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { writeCleanRoomInstallConfig } from "./clean-room-install";

test("clean rooms derive exactly the repository release-age settings on every write", async () => {
  const consumerDir = await mkdtemp(path.join(tmpdir(), "folio-install-policy-"));
  try {
    const repo = Bun.TOML.parse(await readFile(new URL("../bunfig.toml", import.meta.url), "utf8"));
    const policy = Object.fromEntries(
      Object.entries(repo.install).filter(([key]) => key.startsWith("minimumReleaseAge")),
    );
    for (const existing of [
      "",
      '[install]\nminimumReleaseAge = 1\nminimumReleaseAgeExcludes = ["stale"]\n',
    ]) {
      await writeFile(path.join(consumerDir, "bunfig.toml"), existing);
      await writeCleanRoomInstallConfig(consumerDir);
      const written = Bun.TOML.parse(await readFile(path.join(consumerDir, "bunfig.toml"), "utf8"));
      expect(written).toEqual({ install: policy });
    }
  } finally {
    await rm(consumerDir, { recursive: true, force: true });
  }
});

test("the install subprocess receives the policy first, preserves arguments, and propagates failures", async () => {
  const root = await mkdtemp(path.join(tmpdir(), "folio-install-command-"));
  const consumerDir = path.join(root, "consumer with spaces");
  const binDir = path.join(root, "bin");
  try {
    await mkdir(consumerDir);
    await mkdir(binDir);
    await writeFile(
      path.join(binDir, "bun"),
      '#!/bin/sh\ntest -f bunfig.toml || exit 91\ncp bunfig.toml observed.toml\nprintf "%s\\n" "$@" > observed-args\nexit "${INSTALL_TEST_EXIT:-0}"\n',
      { mode: 0o755 },
    );
    const dependencies = [path.join(root, "package with spaces.tgz"), "react@^19"];
    const entry = new URL("./clean-room-install.ts", import.meta.url).href;
    const code = `import { installCleanRoomDependencies } from ${JSON.stringify(entry)}; await installCleanRoomDependencies(${JSON.stringify({ consumerDir, dependencies })});`;
    for (const exitCode of [0, 7]) {
      await rm(path.join(consumerDir, "bunfig.toml"), { force: true });
      const child = Bun.spawn([process.execPath, "-e", code], {
        cwd: root,
        env: {
          ...process.env,
          PATH: `${binDir}${path.delimiter}${process.env.PATH}`,
          INSTALL_TEST_EXIT: String(exitCode),
        },
        stdout: "pipe",
        stderr: "pipe",
      });
      const result = await child.exited;
      expect(result === 0).toBe(exitCode === 0);
      expect(await readFile(path.join(consumerDir, "observed.toml"), "utf8")).toBe(
        await readFile(path.join(consumerDir, "bunfig.toml"), "utf8"),
      );
      expect(
        (await readFile(path.join(consumerDir, "observed-args"), "utf8")).trim().split("\n"),
      ).toEqual(["add", ...dependencies]);
    }
  } finally {
    await rm(root, { recursive: true, force: true });
  }
});

test("dependency consumers explicitly select the isolated linker", async () => {
  const consumerDir = await mkdtemp(path.join(tmpdir(), "folio-isolated-consumer-"));
  try {
    await writeCleanRoomInstallConfig(consumerDir, "isolated");
    const config = Bun.TOML.parse(await readFile(path.join(consumerDir, "bunfig.toml"), "utf8"));
    expect(config.install.linker).toBe("isolated");
  } finally {
    await rm(consumerDir, { recursive: true, force: true });
  }
});
