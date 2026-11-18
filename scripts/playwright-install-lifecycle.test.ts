import { panic } from "better-result";
import { expect, setDefaultTimeout, test } from "bun:test";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

import {
  expectExited,
  runChild,
  subprocessTestTimeout,
} from "../test/subprocess";

setDefaultTimeout(subprocessTestTimeout(4));

const root = fileURLToPath(new URL("../", import.meta.url));
const action: unknown = Bun.YAML.parse(
  await readFile(
    path.join(root, ".github/actions/playwright-deps/action.yml"),
    "utf8",
  ),
);
const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);
const actionScript = (name: string): string => {
  if (!isRecord(action) || !isRecord(action["runs"]))
    panic("Expected action steps");
  const steps = action["runs"]["steps"];
  if (!Array.isArray(steps)) panic("Expected action steps");
  const step = steps.find((value) => isRecord(value) && value["name"] === name);
  if (!isRecord(step) || typeof step["run"] !== "string")
    panic(`Expected ${name} script`);
  return step["run"];
};

const runInstaller = async (fixture: string) => {
  const directory = await mkdtemp(path.join(tmpdir(), "playwright-lifecycle-"));
  const attemptsFile = path.join(directory, "attempts");
  const childFile = path.join(directory, "child");
  const readyFile = path.join(directory, "ready");
  // The test deadline fires only once the parent recorded its child and the
  // child installed its TERM handler. Cold shell startup cannot win the race.
  await writeFile(
    path.join(directory, "sleep"),
    `#!/bin/bash
if [ "$1" = "$ATTEMPT_SECONDS" ]; then
  while [ ! -f "$CHILD_PID" ] || [ ! -f "$CHILD_READY" ]; do /bin/sleep 0.01; done
  exit 0
fi
exec /bin/sleep "$@"
`,
    { mode: 0o755 },
  );
  await writeFile(
    path.join(directory, "bunx"),
    `#!/bin/bash\necho "$$" >> "$INSTALL_ATTEMPTS"\n${fixture}\n`,
    { mode: 0o755 },
  );
  try {
    const {
      exitCode: code,
      stdout,
      stderr,
    } = expectExited(
      runChild(
        [
          "bash",
          "-e",
          "-o",
          "pipefail",
          "-c",
          actionScript("Install Playwright browsers"),
        ],
        {
          cwd: root,
          env: {
            ...process.env,
            PATH: `${directory}:${process.env["PATH"]}`,
            RUNNER_OS_NAME: "macOS",
            BROWSERS: "chromium",
            ATTEMPT_SECONDS: "480",
            KILL_GRACE_SECONDS: "0.1",
            INSTALL_ATTEMPTS: attemptsFile,
            CHILD_PID: childFile,
            CHILD_READY: readyFile,
          },
        },
      ),
    );
    if (!(await Bun.file(attemptsFile).exists()))
      panic(`Fake installer did not start (${code}): ${stdout} ${stderr}`);
    const attempts = (await readFile(attemptsFile, "utf8")).trim().split("\n");
    const descendant = (await Bun.file(childFile).exists())
      ? Number(await readFile(childFile, "utf8"))
      : null;
    let descendantAlive = false;
    if (descendant !== null) {
      const state = expectExited(
        runChild(["ps", "-o", "stat=", "-p", String(descendant)]),
      );
      // A zombie has exited and released its files; PID existence alone
      // cannot distinguish it from an executing, lock-holding descendant.
      descendantAlive =
        state.exitCode === 0 && !state.stdout.trim().startsWith("Z");
    }
    return {
      code,
      stdout,
      stderr,
      attempts: attempts.length,
      descendant,
      descendantAlive,
    };
  } finally {
    // A failing mutant may leave its TERM-ignoring child. Kill only the exact
    // process group started by this fixture, never another test or installer.
    if (await Bun.file(attemptsFile).exists()) {
      for (const pid of (await readFile(attemptsFile, "utf8"))
        .trim()
        .split("\n"))
        expectExited(runChild(["/bin/kill", "-KILL", "--", `-${pid}`]));
    }
    await rm(directory, { recursive: true, force: true });
  }
};

test("timeout completes escalation and never starts a second installer", async () => {
  const result = await runInstaller(`
trap 'exit 143' TERM
bash -c 'trap "" TERM; : > "$CHILD_READY"; while :; do /bin/sleep 1; done' >/dev/null 2>&1 &
echo "$!" > "$CHILD_PID"
wait "$!"
`);
  expect(result.code).toBe(124);
  expect(result.attempts).toBe(1);
  expect(result.stdout).toContain(
    "refusing to start another package installer",
  );
  expect(result.stdout).not.toContain("trying the download once more");
  expect(result.descendant).not.toBeNull();
  // Liveness was observed before the harness's fallback cleanup, so a
  // missing KILL cannot pass merely because the test cleaned up its fixture.
  expect(result.descendantAlive).toBe(false);
});

test.each([
  { fixture: "exit 0", code: 0, attempts: 1 },
  {
    fixture: 'if [ "$(wc -l < "$INSTALL_ATTEMPTS")" -eq 1 ]; then exit 7; fi',
    code: 0,
    attempts: 2,
  },
  { fixture: "exit 7", code: 7, attempts: 2 },
])(
  "completed installers preserve success/failure and one retry ($attempts attempts)",
  async ({ fixture, code, attempts }) => {
    const result = await runInstaller(fixture);
    expect(result.code).toBe(code);
    expect(result.attempts).toBe(attempts);
  },
);

test("the APT setup emits bounded HTTP/HTTPS waits and a finite retry count", async () => {
  const directory = await mkdtemp(
    path.join(tmpdir(), "playwright-apt-config-"),
  );
  const output = path.join(directory, "apt.conf");
  await writeFile(path.join(directory, "mkdir"), "#!/bin/bash\nexit 0\n", {
    mode: 0o755,
  });
  await writeFile(
    path.join(directory, "sudo"),
    '#!/bin/bash\nif [ "$1" = tee ]; then cat > "$APT_CONFIG_OUTPUT"; fi\n',
    { mode: 0o755 },
  );
  try {
    const child = expectExited(
      runChild(
        [
          "bash",
          "-e",
          "-o",
          "pipefail",
          "-c",
          actionScript("Prepare Playwright system packages"),
        ],
        {
          env: {
            ...process.env,
            PATH: `${directory}:${process.env["PATH"]}`,
            APT_CONFIG_OUTPUT: output,
          },
        },
      ),
    );
    expect(child.exitCode).toBe(0);
    const configuration = await readFile(output, "utf8");
    expect(configuration).toContain('Acquire::http::Timeout "30";\n');
    expect(configuration).toContain('Acquire::https::Timeout "30";\n');
    expect(configuration).toContain('Acquire::Retries "1";\n');
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});
