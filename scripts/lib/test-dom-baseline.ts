import { panic } from "better-result";
import { spawnSync } from "node:child_process";
import { readFileSync } from "node:fs";
import path from "node:path";

export const TEST_DOM_BASELINE_PATH =
  ".oxlint-plugins/folio-test-dom.baseline.json";
type DomBaseline = Readonly<Record<string, number>>;

const validateTestDomBaseline = (value: unknown): DomBaseline => {
  if (typeof value !== "object" || value === null || Array.isArray(value))
    panic("DOM baseline must be a JSON object");
  const parsed: Record<string, number> = {};
  for (const [file, count] of Object.entries(value)) {
    if (
      !/^packages\/core\/src\/(?:[\w.-]+\/)*[\w.-]+\.test\.tsx?$/u.test(file) ||
      file.split("/").some((part) => part === "." || part === "..") ||
      typeof count !== "number" ||
      !Number.isSafeInteger(count) ||
      count <= 0
    )
      panic("Invalid DOM baseline entry", { file });
    parsed[file] = count;
  }
  return parsed;
};

export const parseTestDomBaseline = (source: string): DomBaseline =>
  validateTestDomBaseline(JSON.parse(source));

type DomBaselineBase =
  | { type: "committed"; baseline: DomBaseline }
  | {
      type: "bootstrap";
      sourceAtBase: ReadonlyMap<string, Buffer>;
      sourceNow: ReadonlyMap<string, Buffer>;
    };

type CompareTestDomBaselineOptions = {
  current: DomBaseline;
  base: DomBaselineBase;
};

export const compareTestDomBaseline = ({
  current,
  base,
}: CompareTestDomBaselineOptions): string[] => {
  const problems: string[] = [];
  for (const [file, count] of Object.entries(
    validateTestDomBaseline(current),
  )) {
    switch (base.type) {
      case "committed":
        if (count > (base.baseline[file] ?? 0))
          problems.push("DOM baseline allowance grew: " + file);
        break;
      case "bootstrap": {
        const before = base.sourceAtBase.get(file);
        const now = base.sourceNow.get(file);
        if (before === undefined || now === undefined || !before.equals(now))
          problems.push(
            "DOM baseline bootstrap requires byte-identical base source: " +
              file,
          );
        break;
      }
      default: {
        const unsupported: never = base;
        return unsupported;
      }
    }
  }
  return problems;
};

type RunGitOptions = { root: string; args: string[] };

const runGit = ({ root, args }: RunGitOptions): Buffer => {
  const result = spawnSync("git", args, { cwd: root });
  if (result.status !== 0)
    panic("Cannot establish the DOM baseline's actual Git base", {
      args,
      status: result.status,
    });
  return result.stdout;
};

type ReadBaseFileOptions = { root: string; base: string; file: string };

const readBaseFile = ({
  root,
  base,
  file,
}: ReadBaseFileOptions): Buffer | undefined => {
  const listed = runGit({
    root,
    args: ["ls-tree", "--name-only", base, "--", file],
  })
    .toString("utf8")
    .trim();
  if (listed === "") return undefined;
  if (listed !== file) panic("Ambiguous DOM baseline base path", { file });
  return runGit({ root, args: ["show", base + ":" + file] });
};

type CheckTestDomBaselineOptions = {
  root: string;
  current: DomBaseline;
  requestedBase?: string;
};
type CheckTestDomBaselineResult =
  | { type: "committed"; baseSha: string; problems: string[] }
  | {
      type: "bootstrap";
      baseSha: string;
      verifiedSources: string[];
      problems: string[];
    };

export const checkTestDomBaseline = ({
  root,
  current,
  requestedBase = process.env.TOOLCHAIN_BASE,
}: CheckTestDomBaselineOptions): CheckTestDomBaselineResult => {
  const baseSha =
    requestedBase === undefined ||
    requestedBase === "unavailable-toolchain-base"
      ? runGit({ root, args: ["merge-base", "HEAD", "origin/main"] })
          .toString("utf8")
          .trim()
      : requestedBase;
  if (!/^[a-f\d]{40}$/u.test(baseSha))
    panic("Invalid DOM baseline base commit");
  runGit({ root, args: ["cat-file", "-e", baseSha + "^{commit}"] });
  const committed = readBaseFile({
    root,
    base: baseSha,
    file: TEST_DOM_BASELINE_PATH,
  });
  if (committed !== undefined)
    return {
      type: "committed",
      baseSha,
      problems: compareTestDomBaseline({
        current,
        base: {
          type: "committed",
          baseline: parseTestDomBaseline(committed.toString("utf8")),
        },
      }),
    };
  const sourceAtBase = new Map<string, Buffer>();
  const sourceNow = new Map<string, Buffer>();
  for (const file of Object.keys(current)) {
    const source = readBaseFile({ root, base: baseSha, file });
    if (source !== undefined) sourceAtBase.set(file, source);
    sourceNow.set(file, readFileSync(path.join(root, file)));
  }
  return {
    type: "bootstrap",
    baseSha,
    verifiedSources: Object.keys(current),
    problems: compareTestDomBaseline({
      current,
      base: { type: "bootstrap", sourceAtBase, sourceNow },
    }),
  };
};
