import { expect, setDefaultTimeout, test } from "bun:test";
import { execFileSync } from "node:child_process";
import { devNull } from "node:os";
import { basename, dirname, join, resolve } from "node:path";
import { scanTestTimeout } from "../test/scan-timeout";

const root = resolve(import.meta.dir, "..");

setDefaultTimeout(scanTestTimeout(10_000));

test("local env files are ignored at the root and every Vite config directory", () => {
  const tracked = execFileSync("git", ["ls-files", "-z"], { cwd: root, encoding: "utf8" })
    .split("\0")
    .filter(Boolean);
  const configs = tracked.filter((file) =>
    /^vite(?:\.[^.]+)*\.config\.[cm]?[jt]s$/u.test(basename(file)),
  );
  expect(configs.length).toBeGreaterThan(0);

  const directories = new Set([".", ...configs.map((file) => dirname(file))]);
  const names = new Set([".env", ".env.local", ".env.backup", ".env-custom"]);
  for (const mode of ["production", "development", "test", "staging", "custom-mode"]) {
    names.add(`.env.${mode}`);
    names.add(`.env.${mode}.local`);
  }
  const candidates = [...directories].flatMap((directory) =>
    [...names].map((name) => join(directory, name)),
  );
  const samples = [...directories].map((directory) => join(directory, ".env.example"));

  // Query Git without creating env files or relying on a developer's global ignores.
  // --no-index also checks paths if a local env file has already been tracked.
  const ignored = execFileSync(
    "git",
    ["-c", `core.excludesFile=${devNull}`, "check-ignore", "--no-index", "-z", "--stdin"],
    { cwd: root, encoding: "utf8", input: `${[...candidates, ...samples].join("\0")}\0` },
  )
    .split("\0")
    .filter(Boolean);

  expect(new Set(ignored)).toEqual(new Set(candidates));
});
