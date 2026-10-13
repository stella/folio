import { Result, TaggedError, panic } from "better-result";
import { mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import path from "node:path";
import {
  REUSABLE_WORKFLOW_FIXTURES,
  parsePermissionWorkflow,
  workflowFixturePath,
  workflowReference,
  type WorkflowReference,
} from "./lib/reusable-workflow-permissions";

class WorkflowRefreshError extends TaggedError("WorkflowRefreshError")<{
  message: string;
  cause?: unknown;
}> {}

const root = path.resolve(import.meta.dir, "..");
const sources = new Map<string, string>();
const seen = new Set<string>();

const refresh = async (reference: WorkflowReference): Promise<void> => {
  if (seen.has(reference.key)) return;
  seen.add(reference.key);
  const destination = workflowFixturePath(root, reference);
  let source: string;
  switch (reference.type) {
    case "local":
      source = readFileSync(destination, "utf8");
      break;
    case "remote": {
      // Fetch immutable public source only in this maintainer command.
      const url = `https://raw.githubusercontent.com/${reference.repository}/${reference.sha}/.github/workflows/${reference.file}`;
      source = (
        await Result.tryPromise({
          try: async () => {
            const response = await fetch(url, { signal: AbortSignal.timeout(30_000) });
            if (!response.ok)
              throw new WorkflowRefreshError({
                message: `Cannot refresh ${reference.key}: HTTP ${response.status}`,
              });
            return response.text();
          },
          catch: (cause) =>
            new WorkflowRefreshError({ message: `Cannot refresh ${reference.key}`, cause }),
        })
      ).unwrap();
      sources.set(destination, source);
      break;
    }
    default: {
      const unreachable: never = reference;
      return unreachable;
    }
  }
  const workflow = parsePermissionWorkflow(source);
  for (const job of workflow.jobs) {
    if (job.uses === undefined) continue;
    const child = workflowReference(job.uses, reference);
    if (child === undefined)
      panic("Reusable workflow must use a full commit SHA", { uses: job.uses });
    await refresh(child);
  }
};

for (const file of new Bun.Glob(".github/workflows/*.{yml,yaml}").scanSync({ cwd: root })) {
  const workflow = parsePermissionWorkflow(readFileSync(path.join(root, file), "utf8"));
  for (const job of workflow.jobs) {
    if (job.uses === undefined) continue;
    const reference = workflowReference(job.uses);
    if (reference === undefined)
      panic("Reusable workflow must use a full commit SHA", { uses: job.uses });
    await refresh(reference);
  }
}

// Publish only after every source parsed and its transitive calls were fetched.
for (const [file, source] of sources) {
  mkdirSync(path.dirname(file), { recursive: true });
  writeFileSync(file, source);
}
for (const file of new Bun.Glob(`${REUSABLE_WORKFLOW_FIXTURES}/**/*.{yml,yaml}`).scanSync({
  cwd: root,
  dot: true,
})) {
  const absolute = path.join(root, file);
  if (!sources.has(absolute)) rmSync(absolute);
}
// Use the repository format without changing the pinned declarations.
if (sources.size > 0) {
  const formatted = Bun.spawnSync([path.join(root, "node_modules/.bin/oxfmt"), ...sources.keys()], {
    cwd: root,
    stdout: "inherit",
    stderr: "inherit",
  });
  if (formatted.exitCode !== 0)
    throw new WorkflowRefreshError({ message: "Cannot format pinned reusable workflows." });
  for (const [file, source] of sources) {
    if (
      JSON.stringify(Bun.YAML.parse(source)) !==
      JSON.stringify(Bun.YAML.parse(readFileSync(file, "utf8")))
    )
      panic("Formatting changed a pinned workflow declaration", { file });
  }
}
console.log(`Refreshed ${sources.size} pinned reusable workflows.`);
