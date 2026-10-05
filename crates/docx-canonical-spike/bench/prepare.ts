/** Build artifacts before the quiet window; this command takes no measurements. */
import { spawnSync } from "node:child_process";
import { createHash } from "node:crypto";
import { readFileSync, mkdirSync, writeFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { resolve } from "node:path";
import { loadavg } from "node:os";
import binaryen from "binaryen";
import { sourceDigest } from "./artifacts";
import { importPortable, exportPortable } from "./portable";

const root = fileURLToPath(new URL("..", import.meta.url));
const manifest = resolve(root, "Cargo.toml");
const sourceHash = sourceDigest(root);
const command = (binary: string, args: readonly string[]) => {
  if (binary === "nice" && !process.env["CI"] && (loadavg().at(0) ?? Infinity) >= 5)
    throw new TypeError("Cargo build gate: one-minute load must be below 5; no retry.");
  const result = spawnSync(binary, args, { stdio: "inherit" });
  if (result.error) throw result.error;
  if (result.status !== 0) throw new TypeError(`${binary} failed (${result.status}).`);
};
const mode = process.env["RUST_SPIKE_PREPARE_MODE"] ?? "full";
if (mode !== "full" && mode !== "release") throw new TypeError("Unknown preparation mode.");
const portableDirectory = process.env["RUST_SPIKE_PORTABLE_DIR"];
if (portableDirectory) importPortable(root, portableDirectory);
if (mode === "full") {
  command("nice", ["-n", "10", "cargo", "test", "--manifest-path", manifest]);
  command("nice", [
    "-n",
    "10",
    "cargo",
    "build",
    "--manifest-path",
    manifest,
    "--bin",
    "canonical-spike",
  ]);
}
command("nice", [
  "-n",
  "10",
  "cargo",
  "build",
  "--manifest-path",
  manifest,
  "--release",
  "--bin",
  "canonical-spike",
]);
if (!portableDirectory) {
  command("nice", [
    "-n",
    "10",
    "cargo",
    "build",
    "--manifest-path",
    manifest,
    "--release",
    "--target",
    "wasm32-unknown-unknown",
    "--features",
    "wasm",
    "--lib",
  ]);
  command("wasm-bindgen", [
    resolve(root, "target/wasm32-unknown-unknown/release/docx_canonical_spike.wasm"),
    "--target",
    "web",
    "--out-dir",
    resolve(root, "target/wasm-bindgen-release"),
  ]);
}
const wasmBytes = readFileSync(
  resolve(root, "target/wasm-bindgen-release/docx_canonical_spike_bg.wasm"),
);
if (!WebAssembly.validate(wasmBytes)) throw new TypeError("Spike WASM is invalid.");
const module = new WebAssembly.Module(wasmBytes);
if (
  WebAssembly.Module.imports(module).some(({ module: importModule }) =>
    importModule.startsWith("wasi"),
  )
)
  throw new TypeError("Spike WASM must not import WASI.");
const parsed = binaryen.readBinary(wasmBytes);
try {
  if (parsed.getMemoryInfo().shared) throw new TypeError("Spike WASM must not use shared memory.");
} finally {
  parsed.dispose();
}
const glue = readFileSync(
  resolve(root, "target/wasm-bindgen-release/docx_canonical_spike.js"),
  "utf8",
);
if (/\b(?:Worker|SharedArrayBuffer)\b/u.test(glue))
  throw new TypeError("Spike browser glue must stay single-threaded.");
const browserOutput = resolve(root, "target/benchmark-browser");
mkdirSync(browserOutput, { recursive: true });
if (!portableDirectory) {
  const built = await Bun.build({
    entrypoints: [resolve(root, "bench/browser.ts")],
    outdir: browserOutput,
    target: "browser",
    format: "esm",
    minify: true,
    external: ["/wasm/docx_canonical_spike.js"],
  });
  if (!built.success) throw new AggregateError(built.logs, "Browser benchmark bundle failed.");
}
const progress = process.env["RUST_SPIKE_PROGRESS_DIR"];
if (!progress) throw new TypeError("RUST_SPIKE_PROGRESS_DIR is required.");
mkdirSync(progress, { recursive: true });
if (process.env["CI"]) exportPortable(root, resolve(progress, "portable-wasm"));
const artifactPaths = [
  resolve(root, "target/release/canonical-spike"),
  resolve(root, "target/wasm-bindgen-release/docx_canonical_spike.js"),
  resolve(root, "target/wasm-bindgen-release/docx_canonical_spike_bg.wasm"),
  resolve(browserOutput, "browser.js"),
  resolve(root, "bench/run.ts"),
];
const artifacts = artifactPaths.map((path) => ({
  path,
  sha256: createHash("sha256").update(readFileSync(path)).digest("hex"),
}));
const revision = spawnSync("git", ["rev-parse", "HEAD"], {
  cwd: root,
  encoding: "utf8",
}).stdout.trim();
if (sourceHash !== sourceDigest(root))
  throw new TypeError(
    "Source changed during serialized builds; artifacts cannot be stamped ready.",
  );
writeFileSync(
  resolve(progress, "benchmark-artifacts.json"),
  `${JSON.stringify({ status: "built", revision, sourceHash, artifacts, cwd: root, command: [process.execPath, resolve(root, "bench/run.ts")] }, null, 2)}\n`,
);
console.log(
  "Artifacts built; write benchmark-ready.json only after differential and smoke validation.",
);
