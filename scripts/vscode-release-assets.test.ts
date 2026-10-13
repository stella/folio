import { expect, test } from "bun:test";
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

const workflow: unknown = Bun.YAML.parse(
  readFileSync(path.join(import.meta.dir, "../.github/workflows/publish-vscode.yml"), "utf8"),
);
const record = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);
if (!record(workflow) || !record(workflow["jobs"])) throw new Error("Missing publish jobs");
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
  expect(stepsOf("package").filter((step) => step["run"] === "bun run package")).toHaveLength(1);
  const artifact = artifactOptions({
    jobName: "package",
    actionPrefix: "actions/upload-artifact@",
  });
  expect(String(artifact["path"]).trim().split("\n")).toEqual([
    "plugins/vscode/folio.vsix",
    "plugins/vscode/folio.vsix.sha256",
  ]);
  for (const name of Object.keys(jobs).filter((jobName) => jobName !== "package")) {
    expect(jobOf(name)["needs"]).toBe("package");
    const download = artifactOptions({ jobName: name, actionPrefix: "actions/download-artifact@" });
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

test("release upload selects all published tags at the source, including annotated tags", () => {
  const directory = mkdtempSync(path.join(tmpdir(), "folio-release-assets-"));
  const tag = "@stll/folio-core@0.41.0";
  const annotated = "@stll/folio-cli@0.5.1";
  const script = stepsOf("github-release").find(
    (step) => step["name"] === "Upload to published releases at the source commit",
  )?.["run"];
  if (typeof script !== "string") throw new Error("Missing release upload script");
  try {
    writeFileSync(path.join(directory, "git"), '#!/bin/sh\ncat "$TEST_REFS"\n', { mode: 0o755 });
    writeFileSync(
      path.join(directory, "gh"),
      '#!/bin/sh\nif [ "$1" = api ]; then cat "$TEST_RELEASES"; else printf "%s\\n" "$3" >> "$TEST_UPLOADS"; fi\n',
      { mode: 0o755 },
    );
    writeFileSync(
      path.join(directory, "refs"),
      [
        `${SOURCE}\trefs/tags/${tag}`,
        `${OTHER}\trefs/tags/${annotated}`,
        `${SOURCE}\trefs/tags/${annotated}^{}`,
        `${SOURCE}\trefs/tags/draft`,
        `${SOURCE}\trefs/tags/no-release`,
        `${OTHER}\trefs/tags/old-release`,
      ].join("\n") + "\n",
    );
    writeFileSync(path.join(directory, "releases"), `${tag}\n${annotated}\nold-release\n`);
    const result = Bun.spawnSync(["bash", "-c", script], {
      cwd: directory,
      env: {
        ...process.env,
        PATH: `${directory}:${process.env["PATH"] ?? ""}`,
        RUNNER_TEMP: directory,
        REPO: "stella/folio",
        SOURCE_REF: SOURCE,
        TEST_REFS: path.join(directory, "refs"),
        TEST_RELEASES: path.join(directory, "releases"),
        TEST_UPLOADS: path.join(directory, "uploads"),
      },
    });
    expect(result.exitCode).toBe(0);
    expect(readFileSync(path.join(directory, "uploads"), "utf8").trim().split("\n")).toEqual([
      annotated,
      tag,
    ]);
  } finally {
    rmSync(directory, { recursive: true, force: true });
  }
});
