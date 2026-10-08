import { readFileSync } from "node:fs";

/** Check every report, independently of the package selected for extraction. */
export const reportsWithCarriageReturns = (directory: string) =>
  [...new Bun.Glob("**/*.api.md").scanSync({ cwd: directory })]
    .filter((file) => readFileSync(`${directory}/${file}`, "utf8").includes("\r"))
    .sort();
