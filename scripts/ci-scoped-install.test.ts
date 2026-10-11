import { describe, expect, test } from "bun:test";
import {
  linkSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  rmSync,
  symlinkSync,
  writeFileSync,
} from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import {
  INSTALL_SCOPES,
  installedPackageBytes,
  isInstallScope,
  scopeInstall,
} from "./ci-scoped-install";

const root = path.resolve(import.meta.dir, "..");
type JobSteps = { steps: { run?: string; uses?: string }[] };
const checkScopedJobs = (jobs: Record<string, JobSteps>) => {
  for (const scope of Object.keys(INSTALL_SCOPES)) {
    const steps = jobs[scope]?.steps;
    if (steps === undefined) throw new TypeError(`Missing scoped job: ${scope}`);
    expect(
      steps.filter((step) => step.run === `bun scripts/ci-scoped-install.ts ${scope}`),
    ).toHaveLength(1);
    for (const step of steps) {
      expect(step.run ?? "").not.toMatch(/bun\s+(?:install|ci)\b/u);
      expect(step.uses ?? "").not.toContain("setup-bun-cached");
    }
  }
};

describe("scoped CI environments", () => {
  test("package size is invariant under shared links and excludes workspace sources", () => {
    const fixture = mkdtempSync(path.join(tmpdir(), "folio-ci-size-"));
    try {
      const modules = path.join(fixture, "node_modules");
      const source = path.join(fixture, "packages", "workspace");
      mkdirSync(modules);
      mkdirSync(source, { recursive: true });
      writeFileSync(path.join(modules, "payload"), "payload");
      const bytes = installedPackageBytes(fixture);
      expect(bytes).toBe(7);
      linkSync(path.join(modules, "payload"), path.join(modules, "shared"));
      symlinkSync(path.join(modules, "payload"), path.join(modules, "alias"));
      writeFileSync(path.join(source, "source.ts"), "workspace source");
      symlinkSync(source, path.join(modules, "workspace"));
      expect(installedPackageBytes(fixture)).toBe(bytes);
      const alias = path.join(fixture, "root-alias");
      symlinkSync(fixture, alias);
      expect(installedPackageBytes(alias)).toBe(bytes);
    } finally {
      rmSync(fixture, { recursive: true, force: true });
    }
  });
  test("every scoped job uses its environment without restoring the full workspace cache", () => {
    const workflow = Bun.YAML.parse(
      readFileSync(path.join(root, ".github/workflows/ci.yml"), "utf8"),
    );
    checkScopedJobs(workflow.jobs);
    for (const scope of Object.keys(INSTALL_SCOPES)) {
      for (const invalidStep of [
        { run: "bun install --frozen-lockfile" },
        { uses: "stella/.github/actions/setup-bun-cached@fixture" },
      ]) {
        const invalid = structuredClone(workflow.jobs);
        invalid[scope].steps.push(invalidStep);
        expect(() => checkScopedJobs(invalid)).toThrow();
      }
    }
  });

  test("all projections retain source pins, overrides and complete workspace references", () => {
    const original = Bun.JSONC.parse(readFileSync(path.join(root, "bun.lock"), "utf8"));
    const check = (scope: keyof typeof INSTALL_SCOPES) => {
      const { files } = scopeInstall(scope, root);
      const projected = Bun.JSONC.parse(files.get("bun.lock") ?? "");
      const manifest = JSON.parse(files.get("package.json") ?? "");
      expect(projected.overrides).toEqual(original.overrides);
      expect(manifest.devDependencies.binaryen).toBeUndefined();
      expect(projected.workspaces[""].devDependencies).toEqual(manifest.devDependencies);
      for (const [name, entry] of Object.entries(projected.packages)) {
        expect(entry).toEqual(original.packages[name]);
      }
      for (const [dir, workspace] of Object.entries(projected.workspaces)) {
        if (dir === "") continue;
        expect(manifest.workspaces).toContain(dir);
        const encoded =
          files.get(`${dir}/package.json`) ??
          readFileSync(path.join(root, dir, "package.json"), "utf8");
        const source = JSON.parse(encoded);
        expect(workspace).toEqual(
          original.workspaces[dir].devDependencies && scope === "browser-discovery"
            ? Object.fromEntries(
                Object.entries(original.workspaces[dir]).filter(
                  ([key]) => key !== "devDependencies",
                ),
              )
            : original.workspaces[dir],
        );
        expect(source.dependencies).toEqual(original.workspaces[dir].dependencies);
      }
    };
    for (const scope of Object.keys(INSTALL_SCOPES)) {
      if (!isInstallScope(scope)) throw new TypeError(`Unknown install scope: ${scope}`);
      check(scope);
    }
  });

  test("missing declared tool and workspace roots fail before any file changes", () => {
    const fixture = mkdtempSync(path.join(tmpdir(), "folio-ci-scope-"));
    try {
      writeFileSync(
        path.join(fixture, "package.json"),
        JSON.stringify({ name: "fixture", devDependencies: {} }),
      );
      writeFileSync(
        path.join(fixture, "bun.lock"),
        JSON.stringify({ workspaces: { "": { name: "fixture" } }, packages: {} }),
      );
      expect(() => scopeInstall("docs-format", fixture)).toThrow("Undeclared scoped dependency");
      const manifest = JSON.parse(readFileSync(path.join(root, "package.json"), "utf8"));
      writeFileSync(path.join(fixture, "package.json"), JSON.stringify(manifest));
      expect(() => scopeInstall("browser-discovery", fixture)).toThrow("Unknown workspace");
    } finally {
      rmSync(fixture, { recursive: true, force: true });
    }
  });
});
