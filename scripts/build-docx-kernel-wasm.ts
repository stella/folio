#!/usr/bin/env bun

import binaryen from "binaryen";
import { panic } from "better-result";
import { createHash, timingSafeEqual } from "node:crypto";
import {
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  readdirSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { brotliCompressSync, constants as zlibConstants } from "node:zlib";
import { $ } from "bun";

const mode = process.argv[2];
if (mode !== "--check" && mode !== "--write") {
  panic("usage: bun scripts/build-docx-kernel-wasm.ts <--check|--write>");
}

const repoRoot = path.resolve(import.meta.dir, "..");
const generatedDir = path.join(repoRoot, "packages", "docx-core", "src", "generated");
const temporaryDir = mkdtempSync(path.join(tmpdir(), "folio-docx-kernel-wasm-"));
const artifact = path.join(
  repoRoot,
  "target",
  "wasm32-unknown-unknown",
  "release",
  "stella_docx_kernel.wasm",
);
const generatedFiles = [
  "docx_kernel.js",
  "docx_kernel.d.ts",
  "docx_kernel_bg.wasm",
  "docx_kernel_bg.wasm.d.ts",
  "docx_kernel.inputs.sha256",
] as const;
const wasmFile = "docx_kernel_bg.wasm";
const canonicalArtifactPlatform = process.platform === "linux" && process.arch === "x64";
const maximumWasmBytes = 191 * 1024;
const maximumBrotliBytes = 90 * 1024;

const sourceFiles = readdirSync(path.join(repoRoot, "crates", "docx-kernel", "src"), {
  recursive: true,
})
  .filter((file): file is string => typeof file === "string" && file.endsWith(".rs"))
  .map((file) => path.join(repoRoot, "crates", "docx-kernel", "src", file));
const generationInputs = [
  path.join(repoRoot, "Cargo.lock"),
  path.join(repoRoot, "Cargo.toml"),
  path.join(repoRoot, "rust-toolchain.toml"),
  path.join(repoRoot, "crates", "docx-kernel", "Cargo.toml"),
  path.join(repoRoot, "scripts", "build-docx-kernel-wasm.ts"),
  ...sourceFiles,
].sort();

const generationInputDigest = (): string => {
  const digest = createHash("sha256");
  for (const file of generationInputs) {
    digest.update(path.relative(repoRoot, file));
    digest.update("\0");
    digest.update(readFileSync(file));
    digest.update("\0");
  }
  return `${digest.digest("hex")}\n`;
};

const capture = async (command: string, arguments_: string[]): Promise<string> => {
  const result = await $`${command} ${arguments_}`.cwd(repoRoot).quiet();
  return result.text().trim();
};

const cargoMetadata = JSON.parse(
  await capture("cargo", ["metadata", "--locked", "--format-version", "1", "--features", "wasm"]),
) as {
  packages: { name: string; version: string }[];
};
const wasmBindgenPackages = cargoMetadata.packages.filter(({ name }) => name === "wasm-bindgen");
if (wasmBindgenPackages.length !== 1) {
  panic(
    `Expected exactly one wasm-bindgen package; found ${wasmBindgenPackages.map(({ version }) => version).join(", ") || "none"}`,
  );
}
const wasmBindgenVersion =
  wasmBindgenPackages.at(0)?.version ?? panic("Could not resolve the pinned wasm-bindgen runtime");
const installedWasmBindgen = await capture("wasm-bindgen", ["--version"]);
const expectedWasmBindgen = `wasm-bindgen ${wasmBindgenVersion}`;
if (installedWasmBindgen !== expectedWasmBindgen) {
  panic(`Expected ${expectedWasmBindgen}; found ${installedWasmBindgen}`);
}

try {
  await $`cargo build --locked --release --target wasm32-unknown-unknown --features wasm -p stella-docx-kernel`.cwd(
    repoRoot,
  );
  await $`wasm-bindgen ${artifact} --target web --remove-name-section --remove-producers-section --out-dir ${temporaryDir} --out-name docx_kernel`.cwd(
    repoRoot,
  );

  const generatedJavaScriptPath = path.join(temporaryDir, "docx_kernel.js");
  const generatedJavaScript = readFileSync(generatedJavaScriptPath, "utf8");
  const packageRelativeJavaScript = generatedJavaScript.replace(
    "new URL('docx_kernel_bg.wasm', import.meta.url)",
    "new URL('./docx_kernel_bg.wasm', import.meta.url)",
  );
  if (packageRelativeJavaScript === generatedJavaScript) {
    panic("Could not normalize wasm-bindgen's package-relative WebAssembly URL");
  }
  writeFileSync(
    generatedJavaScriptPath,
    `// @ts-nocheck -- generated and verified by build-docx-kernel-wasm.ts\n${packageRelativeJavaScript}`,
  );

  const wasmPath = path.join(temporaryDir, "docx_kernel_bg.wasm");
  const wasmModule = binaryen.readBinary(readFileSync(wasmPath));
  binaryen.setOptimizeLevel(3);
  binaryen.setShrinkLevel(2);
  wasmModule.setFeatures(
    wasmModule.getFeatures() | binaryen.Features.BulkMemory | binaryen.Features.BulkMemoryOpt,
  );
  wasmModule.optimize();
  wasmModule.optimize();
  if (wasmModule.getMemoryInfo().shared) {
    panic("DOCX kernel WebAssembly must not require shared memory or threads");
  }
  writeFileSync(wasmPath, wasmModule.emitBinary());
  wasmModule.dispose();
  writeFileSync(path.join(temporaryDir, "docx_kernel.inputs.sha256"), generationInputDigest());

  const wasmBytes = readFileSync(wasmPath);
  const brotliBytes = brotliCompressSync(wasmBytes, {
    params: { [zlibConstants.BROTLI_PARAM_QUALITY]: 9 },
  });
  if (!WebAssembly.validate(wasmBytes)) {
    panic("Generated DOCX kernel artifact is not valid WebAssembly");
  }
  const compiledModule = new WebAssembly.Module(wasmBytes);
  const wasiImports = WebAssembly.Module.imports(compiledModule).filter(({ module }) =>
    module.startsWith("wasi"),
  );
  if (wasiImports.length > 0) {
    panic(
      `DOCX kernel WebAssembly must not import WASI: ${wasiImports.map(({ module, name }) => `${module}.${name}`).join(", ")}`,
    );
  }
  if (wasmBytes.byteLength > maximumWasmBytes) {
    panic(
      `DOCX kernel WebAssembly is ${wasmBytes.byteLength} bytes; budget is ${maximumWasmBytes}`,
    );
  }
  if (brotliBytes.byteLength > maximumBrotliBytes) {
    panic(
      `DOCX kernel Brotli size is ${brotliBytes.byteLength} bytes; budget is ${maximumBrotliBytes}`,
    );
  }

  if (mode === "--write") {
    mkdirSync(generatedDir, { recursive: true });
    for (const file of generatedFiles) {
      writeFileSync(path.join(generatedDir, file), readFileSync(path.join(temporaryDir, file)));
    }
  } else {
    const comparableFiles = canonicalArtifactPlatform
      ? generatedFiles
      : generatedFiles.filter((file) => file !== wasmFile);
    const stale = comparableFiles.filter((file) => {
      const expectedPath = path.join(generatedDir, file);
      if (!existsSync(expectedPath)) {
        return true;
      }
      const expected = createHash("sha256").update(readFileSync(expectedPath)).digest();
      const actual = createHash("sha256")
        .update(readFileSync(path.join(temporaryDir, file)))
        .digest();
      return !timingSafeEqual(expected, actual);
    });
    if (stale.length > 0) {
      panic(
        `Generated DOCX kernel artifacts are stale: ${stale.join(", ")}. Run bun --filter @stll/docx-core wasm:generate.`,
      );
    }

    if (!canonicalArtifactPlatform) {
      const committedWasmPath = path.join(generatedDir, wasmFile);
      if (
        !existsSync(committedWasmPath) ||
        !WebAssembly.validate(readFileSync(committedWasmPath))
      ) {
        panic("Committed DOCX kernel artifact is not valid WebAssembly");
      }
    }
  }

  const generatedFileSet: ReadonlySet<string> = new Set(generatedFiles);
  const unexpected = readdirSync(temporaryDir).filter((file) => !generatedFileSet.has(file));
  if (unexpected.length > 0) {
    panic(`Unexpected wasm-bindgen artifacts: ${unexpected.join(", ")}`);
  }
  console.log(
    `DOCX kernel WebAssembly: ${wasmBytes.byteLength} bytes raw, ${brotliBytes.byteLength} bytes Brotli`,
  );
} finally {
  rmSync(temporaryDir, { recursive: true, force: true });
}
