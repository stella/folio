// Scratch runner (not committed): replay every flow, one line each.
import { readdirSync, readFileSync, writeFileSync } from "node:fs";

import { runFlowFile, vacuousSteps } from "./support/fuzz.ts";

const [out] = process.argv.slice(2);
const dir = new URL("./flows/", import.meta.url).pathname;
const lines: string[] = [];
for (const file of readdirSync(dir)
  .filter((name) => name.endsWith(".json"))
  .toSorted()) {
  const flow = JSON.parse(readFileSync(dir + file, "utf8"));
  try {
    const run = await runFlowFile(flow);
    lines.push(`${file}\tOK\t${JSON.stringify(vacuousSteps(run))}`);
  } catch (error) {
    const message = String(error instanceof Error ? error.message : error);
    const step = /failed at step (\d+)/u.exec(message)?.[1] ?? "?";
    const detail = message
      .split("\n")
      .filter((line) => !/^\s+\d+\. /u.test(line))
      .slice(1, 3)
      .join(" | ");
    lines.push(`${file}\tFAIL\tstep ${step}\t${detail.slice(0, 200)}`);
  }
}
writeFileSync(out ?? "/dev/stdout", `${lines.join("\n")}\n`);
