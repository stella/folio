import { $ } from "bun";
import { execFileSync, execSync, spawnSync } from "node:child_process";

const consumerDir = "/tmp/consumer";
await $`bun add react`.cwd(consumerDir);
await $`bun install --frozen-lockfile`.cwd(consumerDir);
Bun.spawn(["bun", "add", "react"], { cwd: consumerDir });
Bun.spawnSync([process.execPath, "install"], { cwd: consumerDir });
spawnSync("bun", ["add", "react"], { cwd: consumerDir });
execFileSync("/usr/local/bin/bun", ["install"], { cwd: consumerDir });
await $`cd ${consumerDir} && bun --no-progress install`;
execSync("bun install", { cwd: consumerDir });
Bun.spawn({ cmd: ["bun", "--no-progress", "add", "react"], cwd: consumerDir });
