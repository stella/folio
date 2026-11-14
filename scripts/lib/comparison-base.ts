import { spawnSync } from "node:child_process";

/**
 * The workflow's TOOLCHAIN_BASE when its event names no base commit, as on a
 * hand-started run.
 */
export const UNNAMED_BASE = "unavailable-toolchain-base";

const MAIN_REF = "origin/main";

type ResolveComparisonBaseOptions = {
  root: string;
  requested?: string | undefined;
};

/**
 * The commit a check compares this checkout with. A base the event names is
 * kept. Without one, the base is where HEAD left main: the commit a pull
 * request of the same branch builds on, and HEAD itself on main.
 */
export const resolveComparisonBase = ({
  root,
  requested,
}: ResolveComparisonBaseOptions): string => {
  if (requested !== undefined && requested !== "" && requested !== UNNAMED_BASE)
    return requested;
  const result = spawnSync("git", ["merge-base", "HEAD", MAIN_REF], {
    cwd: root,
    encoding: "utf8",
  });
  const base = result.stdout.trim();
  if (result.status !== 0 || !/^[a-f\d]{40}$/u.test(base))
    throw new Error(
      `Cannot find where HEAD left ${MAIN_REF}: git merge-base exited ${String(result.status)}: ${result.stderr.trim()}`,
    );
  return base;
};

// A workflow step's output line: `base=<commit>`.
if (import.meta.main) {
  console.log(
    `base=${resolveComparisonBase({
      root: process.cwd(),
      requested: process.env["TOOLCHAIN_BASE"],
    })}`,
  );
}
