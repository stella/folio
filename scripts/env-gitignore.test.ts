import { expect, setDefaultTimeout, test } from "bun:test";
import { devNull } from "node:os";
import { basename, dirname, join, resolve } from "node:path";

import { scanTestTimeout } from "../test/scan-timeout";
import {
  expectExited,
  runChild,
  subprocessTestTimeout,
} from "../test/subprocess";

const root = resolve(import.meta.dir, "..");

setDefaultTimeout(scanTestTimeout(10_000));

/** Git's output for `args`; git must exit cleanly. */
const git = (args: string[], input?: string): string => {
  const run = expectExited(runChild(["git", ...args], { cwd: root, input }));
  expect(run.exitCode).toBe(0);
  return run.stdout;
};

test(
  "local env files are ignored at the root and every Vite config directory",
  () => {
    const tracked = git(["ls-files", "-z"]).split("\0").filter(Boolean);
    const configs = tracked.filter((file) =>
      /^vite(?:\.[^.]+)*\.config\.[cm]?[jt]s$/u.test(basename(file)),
    );
    expect(configs.length).toBeGreaterThan(0);

    const directories = new Set([".", ...configs.map((file) => dirname(file))]);
    const names = new Set([".env", ".env.local", ".env.backup", ".env-custom"]);
    for (const mode of [
      "production",
      "development",
      "test",
      "staging",
      "custom-mode",
    ]) {
      names.add(`.env.${mode}`);
      names.add(`.env.${mode}.local`);
    }
    const candidates = [...directories].flatMap((directory) =>
      [...names].map((name) => join(directory, name)),
    );
    const samples = [...directories].map((directory) =>
      join(directory, ".env.example"),
    );

    // Query Git without creating env files or relying on a developer's global ignores.
    // --no-index also checks paths if a local env file has already been tracked.
    const ignored = git(
      [
        "-c",
        `core.excludesFile=${devNull}`,
        "check-ignore",
        "--no-index",
        "-z",
        "--stdin",
      ],
      `${[...candidates, ...samples].join("\0")}\0`,
    )
      .split("\0")
      .filter(Boolean);

    expect(new Set(ignored)).toEqual(new Set(candidates));

    // An ignore rule does not untrack a file Git already tracks.
    expect(
      tracked.filter(
        (file) =>
          /^\.env(?:[.-]|$)/u.test(basename(file)) &&
          basename(file) !== ".env.example",
      ),
    ).toEqual([]);
  },
  subprocessTestTimeout(2),
);
