import { expect, test } from "bun:test";
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

const workflow: unknown = Bun.YAML.parse(
  readFileSync(
    path.join(import.meta.dir, "../.github/workflows/publish-vscode.yml"),
    "utf8",
  ),
);
const record = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);
if (!record(workflow) || !record(workflow["jobs"]))
  throw new Error("Missing publish jobs");
const jobs = workflow["jobs"];
const jobOf = (name: string) => {
  const job = jobs[name];
  if (!record(job)) throw new Error(`Missing publish job ${name}`);
  return job;
};
const stepsOf = (name: string) => {
  const steps = jobOf(name)["steps"];
  if (!Array.isArray(steps)) throw new Error(`Missing steps in ${name}`);
  return steps.filter(record);
};
type ArtifactOptionsArgs = { jobName: string; actionPrefix: string };
const artifactOptions = ({ jobName, actionPrefix }: ArtifactOptionsArgs) => {
  const step = stepsOf(jobName).find((candidate) =>
    String(candidate["uses"] ?? "").startsWith(actionPrefix),
  );
  if (step === undefined || !record(step["with"]))
    throw new Error(`Missing artifact step in ${jobName}`);
  return step["with"];
};
const SOURCE = "1234567890abcdef1234567890abcdef12345678";
const OTHER = "abcdef1234567890abcdef1234567890abcdef12";

test("each destination consumes the one packaged VSIX and verifies its checksum", () => {
  expect(
    stepsOf("package").filter((step) => step["run"] === "bun run package"),
  ).toHaveLength(1);
  const artifact = artifactOptions({
    jobName: "package",
    actionPrefix: "actions/upload-artifact@",
  });
  expect(String(artifact["path"]).trim().split("\n")).toEqual([
    "plugins/vscode/folio.vsix",
    "plugins/vscode/folio.vsix.sha256",
  ]);
  for (const name of Object.keys(jobs).filter(
    (jobName) => jobName !== "package",
  )) {
    expect(jobOf(name)["needs"]).toBe("package");
    const download = artifactOptions({
      jobName: name,
      actionPrefix: "actions/download-artifact@",
    });
    expect(download["name"]).toBe(artifact["name"]);
    const commands = stepsOf(name)
      .map((step) => step["run"] ?? "")
      .join("\n");
    expect(commands).toContain("sha256sum --check folio.vsix.sha256");
    expect(commands).not.toContain("bun run package");
  }
  expect(jobOf("github-release")["if"]).toBe("inputs.publish");
  expect(jobOf("github-release")["permissions"]).toEqual({ contents: "write" });
});

type ReleaseRun = { version: string; existing: string | null };
const runRelease = ({
  version,
  existing,
}: ReleaseRun): { exitCode: number; calls: string[] } => {
  const directory = mkdtempSync(path.join(tmpdir(), "folio-release-assets-"));
  const script = stepsOf("github-release").find(
    (step) => step["name"] === "Publish the extension release",
  )?.["run"];
  if (typeof script !== "string")
    throw new Error("Missing extension release script");
  try {
    writeFileSync(
      path.join(directory, "gh"),
      [
        "#!/bin/sh",
        'printf "%s\\n" "$*" >> "$TEST_CALLS"',
        'if [ "$1 $2" = "release view" ]; then [ -n "$TEST_EXISTING" ]; exit $?; fi',
        'if [ "$1" = api ]; then printf "%s\\n" "$TEST_EXISTING"; fi',
        "exit 0",
      ].join("\n") + "\n",
      { mode: 0o755 },
    );
    const result = Bun.spawnSync(["bash", "-c", script], {
      cwd: directory,
      env: {
        ...process.env,
        PATH: `${directory}:${process.env["PATH"] ?? ""}`,
        REPO: "stella/folio",
        SOURCE_REF: SOURCE,
        VERSION: version,
        TEST_EXISTING: existing ?? "",
        TEST_CALLS: path.join(directory, "calls"),
      },
    });
    let calls: string[] = [];
    try {
      calls = readFileSync(path.join(directory, "calls"), "utf8")
        .trim()
        .split("\n");
    } catch {
      calls = [];
    }
    return { exitCode: result.exitCode, calls };
  } finally {
    rmSync(directory, { recursive: true, force: true });
  }
};

test("the extension gets its own release at the source commit and never touches package releases", () => {
  const { exitCode, calls } = runRelease({ version: "0.5.3", existing: null });
  expect(exitCode).toBe(0);
  const create = calls.find((call) => call.startsWith("release create "));
  expect(create).toStartWith(
    "release create folio-vscode@0.5.3 folio.vsix folio.vsix.sha256 ",
  );
  expect(create).toContain(`--target ${SOURCE}`);
  expect(create).toContain("--latest=false");
  expect(calls.filter((call) => call.startsWith("release upload "))).toEqual(
    [],
  );
  expect(calls.every((call) => !call.includes("@stll/"))).toBe(true);
});

test("a retry from the same commit replaces the files of the existing extension release", () => {
  const { exitCode, calls } = runRelease({
    version: "0.5.3",
    existing: SOURCE,
  });
  expect(exitCode).toBe(0);
  expect(calls.filter((call) => call.startsWith("release upload "))).toEqual([
    "release upload folio-vscode@0.5.3 folio.vsix folio.vsix.sha256 --repo stella/folio --clobber",
  ]);
  expect(calls.some((call) => call.startsWith("release create "))).toBe(false);
});

test("a later commit at an unchanged version leaves the existing extension release alone", () => {
  const { exitCode, calls } = runRelease({ version: "0.5.3", existing: OTHER });
  expect(exitCode).toBe(0);
  expect(calls.some((call) => call.startsWith("release upload "))).toBe(false);
  expect(calls.some((call) => call.startsWith("release create "))).toBe(false);
});

test("an unreleased extension version publishes nothing", () => {
  const { exitCode, calls } = runRelease({ version: "0.0.0", existing: null });
  expect(exitCode).toBe(0);
  expect(calls).toEqual([]);
});
