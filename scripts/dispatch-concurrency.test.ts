import { expect, test } from "bun:test";
import { readdirSync, readFileSync } from "node:fs";
import path from "node:path";

const WORKFLOWS = path.resolve(import.meta.dir, "../.github/workflows");

/**
 * Workflows a hand-started round dispatches on one fixed commit. A push, a
 * schedule or another dispatch must never cancel or replace such a run, so a
 * dispatch gets a concurrency group of its own, keyed on its run id; every
 * other event keeps the group it had.
 */
const ROUND_GROUPS: Readonly<Record<string, string>> = {
  "ci.yml": "format('ci-{0}-{1}', github.workflow, github.ref)",
  "continuous-fuzz.yml": "'continuous-fuzz'",
  "nightly-long-fuzz.yml": "format('nightly-long-fuzz-{0}', github.ref)",
  "nightly-public-corpus-flows.yml":
    "format('nightly-public-corpus-flows-{0}', github.ref)",
  "nightly-swarm-fuzz.yml": "format('nightly-swarm-fuzz-{0}', github.ref)",
};

/** Dispatchable workflows that serialize on purpose, and why. */
const SERIALIZED: Readonly<Record<string, string>> = {
  "benchmarks.yml": "runs on one ref queue so two never measure side by side",
  "bun-binary-cache.yml": "one writer for the shared Bun binary cache",
  "canonical-wasm.yml": "projection artifacts follow the newest commit",
  "publish-crate.yml": "one crate release at a time",
  "publish-vscode.yml": "one extension release at a time",
  "publish.yml": "one package release at a time",
  "test-timing-refresh.yml": "one timing refresh pull request at a time",
};

type Workflow = { on?: unknown; concurrency?: unknown };

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const workflows = readdirSync(WORKFLOWS)
  .filter((file) => file.endsWith(".yml"))
  .map((file) => {
    const parsed: unknown = Bun.YAML.parse(
      readFileSync(path.join(WORKFLOWS, file), "utf8"),
    );
    if (!isRecord(parsed)) throw new TypeError(`${file} is not a mapping`);
    const workflow: Workflow = parsed;
    return { file, workflow };
  });

const dispatchable = ({ on }: Workflow): boolean =>
  isRecord(on) && "workflow_dispatch" in on;

test("a dispatched round run has a concurrency group of its own", () => {
  for (const [file, otherEvents] of Object.entries(ROUND_GROUPS)) {
    const entry = workflows.find((candidate) => candidate.file === file);
    if (entry === undefined) throw new Error(`Missing workflow ${file}`);
    expect(dispatchable(entry.workflow)).toBe(true);
    const concurrency = entry.workflow.concurrency;
    if (!isRecord(concurrency)) throw new TypeError(`${file} has no group`);
    const name = file.replace(/\.yml$/u, "");
    expect([file, concurrency["group"]]).toEqual([
      file,
      `\${{ github.event_name == 'workflow_dispatch' && format('${name}-dispatch-{0}', github.run_id) || ${otherEvents} }}`,
    ]);
  }
});

// A new dispatchable workflow with a shared group has to choose: a round run
// that nothing cancels, or a stated reason to serialize.
test("every other dispatchable shared group states why it serializes", () => {
  const shared = workflows
    .filter(
      ({ file, workflow }) =>
        dispatchable(workflow) &&
        workflow.concurrency !== undefined &&
        !(file in ROUND_GROUPS),
    )
    .map(({ file }) => file);
  expect(shared.toSorted()).toEqual(Object.keys(SERIALIZED).toSorted());
});
