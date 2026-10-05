/** A readiness stamp means the built browser/native fixtures were checked, never merely compiled. */
import { spawnSync } from "node:child_process";
import { readFileSync, writeFileSync } from "node:fs";
import { resolve } from "node:path";
import { createHash } from "node:crypto";
import { fileURLToPath } from "node:url";
import { loadavg } from "node:os";

if (!process.env["CI"] && (loadavg().at(0) ?? Infinity) >= 8)
  throw new TypeError("Targeted validation requires one-minute load below 8; no retry.");
const progress = process.env["RUST_SPIKE_PROGRESS_DIR"];
if (!progress) throw new TypeError("RUST_SPIKE_PROGRESS_DIR is required.");
const artifacts: unknown = JSON.parse(
  readFileSync(resolve(progress, "benchmark-artifacts.json"), "utf8"),
);
if (
  typeof artifacts !== "object" ||
  artifacts === null ||
  !("status" in artifacts) ||
  artifacts.status !== "built" ||
  !("artifacts" in artifacts) ||
  !Array.isArray(artifacts.artifacts)
)
  throw new TypeError("Missing built artifact manifest.");
for (const artifact of artifacts.artifacts) {
  if (typeof artifact?.path !== "string" || typeof artifact.sha256 !== "string")
    throw new TypeError("Malformed artifact record.");
  if (createHash("sha256").update(readFileSync(artifact.path)).digest("hex") !== artifact.sha256)
    throw new TypeError("Artifact changed since build.");
}
const root = fileURLToPath(new URL("..", import.meta.url));
const commands = [
  [resolve(root, "tests/prepare-comment-oracle.ts")],
  [
    "test",
    resolve(root, "tests/native-differential.test.ts"),
    resolve(root, "tests/harness-codec.test.ts"),
    resolve(root, "tests/comment-differential.test.ts"),
  ],
  [resolve(root, "bench/run.ts")],
];
for (const args of commands) {
  const result = spawnSync(process.execPath, args, {
    stdio: "inherit",
    env: {
      ...process.env,
      RUST_SPIKE_BENCH_MODE: "validate",
      RUST_SPIKE_NATIVE_BINARY: resolve(root, "target/release/canonical-spike"),
    },
  });
  if (result.error) throw result.error;
  if (result.status !== 0) throw new TypeError(`Artifact validation failed (${result.status}).`);
}
writeFileSync(
  resolve(progress, "benchmark-ready.json"),
  `${JSON.stringify({ ...artifacts, status: "validated", validatedAt: `${new Date().toLocaleString("sv-SE", { timeZone: "Europe/Prague" })} CEST`, validation: "Native tagged differential/codec/comment oracles and all browser/native benchmark fixtures; no timing samples." }, null, 2)}\n`,
);
