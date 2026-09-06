import { expect, test } from "bun:test";
import { cp, mkdir, mkdtemp, readFile, rm, symlink, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

test("version changes synchronize every inherited crate and reach a fixed point", async () => {
  const root = await mkdtemp(path.join(tmpdir(), "folio-cargo-version-"));
  try {
    await mkdir(path.join(root, "scripts"));
    await mkdir(path.join(root, "packages/docx-core"), { recursive: true });
    await cp(
      import.meta.path.replace(".test.ts", ".ts"),
      path.join(root, "scripts/sync-docx-kernel-version.ts"),
    );
    await symlink(
      path.resolve(import.meta.dir, "../node_modules"),
      path.join(root, "node_modules"),
      "dir",
    );
    await writeFile(
      path.join(root, "Cargo.toml"),
      '[workspace]\nmembers = ["crates/*"]\nresolver = "2"\n\n[workspace.package]\nversion = "1.0.0"\n',
    );
    const names = [
      "stella-docx-kernel",
      "stella-text-shaper",
      "future-member",
      "independent-member",
    ];
    for (const name of names) {
      const crate = path.join(root, "crates", name);
      await mkdir(path.join(crate, "src"), { recursive: true });
      await writeFile(path.join(crate, "src/lib.rs"), "pub const VALUE: u8 = 1;\n");
      await writeFile(
        path.join(crate, "Cargo.toml"),
        `[package]\nname = "${name}"\n${name === "independent-member" ? 'version = "0.5.0"' : "version.workspace = true"}\nedition = "2021"\n`,
      );
    }
    const run = (command: string[]) =>
      Bun.spawnSync(command, { cwd: root, env: { ...process.env, CARGO_NET_OFFLINE: "true" } });
    const synchronize = () =>
      run([process.execPath, "scripts/sync-docx-kernel-version.ts", "--write"]);
    const check = () => run([process.execPath, "scripts/sync-docx-kernel-version.ts"]);
    expect(run(["cargo", "generate-lockfile"]).exitCode).toBe(0);
    const initialLock = await readFile(path.join(root, "Cargo.lock"), "utf8");
    for (const version of ["2.0.0", "2.0.1", "1.0.0"]) {
      await writeFile(
        path.join(root, "packages/docx-core/package.json"),
        JSON.stringify({ version }),
      );
      expect(check().exitCode).not.toBe(0);
      const result = synchronize();
      expect(result.exitCode).toBe(0);
      const lock = await readFile(path.join(root, "Cargo.lock"), "utf8");
      for (const name of names) {
        expect(lock).toContain(
          `name = "${name}"\nversion = "${name === "independent-member" ? "0.5.0" : version}"`,
        );
      }
      expect(check().exitCode).toBe(0);
      expect(synchronize().exitCode).toBe(0);
      expect(await readFile(path.join(root, "Cargo.lock"), "utf8")).toBe(lock);
    }
    expect(await readFile(path.join(root, "Cargo.lock"), "utf8")).toBe(initialLock);
    const lockPath = path.join(root, "Cargo.lock");
    await writeFile(
      lockPath,
      initialLock.replace(
        'name = "stella-text-shaper"\nversion = "1.0.0"',
        'name = "stella-text-shaper"\nversion = "0.9.0"',
      ),
    );
    expect(check().exitCode).not.toBe(0);
    expect(synchronize().exitCode).toBe(0);
    expect(await readFile(lockPath, "utf8")).toBe(initialLock);
  } finally {
    await rm(root, { recursive: true, force: true });
  }
}, 30_000);
