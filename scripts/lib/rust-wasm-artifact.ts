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

/**
 * One committed WebAssembly artifact built from one crate.
 *
 * Every artifact goes through this same pipeline, so a rule that holds for one
 * (no WASI imports, no shared memory, a size budget, byte-for-byte staleness
 * against what is committed) holds for all of them. A second copy of the
 * pipeline would let the artifacts drift in what they are allowed to be.
 */
export type RustWasmArtifact = {
  /** Names the artifact in diagnostics: "DOCX kernel", "text shaper". */
  readonly label: string;
  readonly crate: string;
  /** Crate directory relative to the repository root. */
  readonly crateDirectory: string;
  /** Cargo's output file stem, which is the crate name with underscores. */
  readonly cargoArtifact: string;
  /** wasm-bindgen's `--out-name`, and so the stem of every generated file. */
  readonly outName: string;
  /** Where the generated files are committed, relative to the root. */
  readonly generatedDirectory: string;
  /** What a developer runs to refresh the artifact. */
  readonly regenerateCommand: string;
  readonly maximumWasmBytes: number;
  readonly maximumBrotliBytes: number;
  /** The entry script, relative to the root; itself an input to the digest. */
  readonly buildScript: string;
};

export type BuildMode = "--check" | "--write";

export const buildModeFrom = (argument: string | undefined, usage: string): BuildMode => {
  if (argument !== "--check" && argument !== "--write") {
    panic(`usage: ${usage} <--check|--write>`);
  }
  return argument;
};

const repoRoot = path.resolve(import.meta.dir, "..", "..");

// The committed WebAssembly bytes are only comparable against a build from the
// same platform: another host's codegen is a legitimate difference, not drift.
// Everything else wasm-bindgen emits is platform-independent.
const canonicalArtifactPlatform = process.platform === "linux" && process.arch === "x64";

const capture = async (command: string, arguments_: string[]): Promise<string> => {
  const result = await $`${command} ${arguments_}`.cwd(repoRoot).quiet();
  return result.text().trim();
};

const generationInputsOf = (artifact: RustWasmArtifact): string[] => {
  const sourceDirectory = path.join(repoRoot, artifact.crateDirectory, "src");
  const sourceFiles = readdirSync(sourceDirectory, { recursive: true })
    .filter((file): file is string => typeof file === "string" && file.endsWith(".rs"))
    .map((file) => path.join(sourceDirectory, file));
  return [
    path.join(repoRoot, "Cargo.lock"),
    path.join(repoRoot, "Cargo.toml"),
    path.join(repoRoot, "rust-toolchain.toml"),
    path.join(repoRoot, artifact.crateDirectory, "Cargo.toml"),
    path.join(repoRoot, "scripts", "lib", "rust-wasm-artifact.ts"),
    path.join(repoRoot, artifact.buildScript),
    ...sourceFiles,
  ].sort();
};

const generationInputDigest = (artifact: RustWasmArtifact): string => {
  const digest = createHash("sha256");
  for (const file of generationInputsOf(artifact)) {
    digest.update(path.relative(repoRoot, file));
    digest.update("\0");
    digest.update(readFileSync(file));
    digest.update("\0");
  }
  return `${digest.digest("hex")}\n`;
};

const assertPinnedWasmBindgen = async (): Promise<void> => {
  const cargoMetadata = JSON.parse(
    await capture("cargo", ["metadata", "--locked", "--format-version", "1", "--features", "wasm"]),
  ) as { packages: { name: string; version: string }[] };
  const wasmBindgenPackages = cargoMetadata.packages.filter(({ name }) => name === "wasm-bindgen");
  if (wasmBindgenPackages.length !== 1) {
    panic(
      `Expected exactly one wasm-bindgen package; found ${wasmBindgenPackages.map(({ version }) => version).join(", ") || "none"}`,
    );
  }
  const wasmBindgenVersion =
    wasmBindgenPackages.at(0)?.version ??
    panic("Could not resolve the pinned wasm-bindgen runtime");
  const installedWasmBindgen = await capture("wasm-bindgen", ["--version"]);
  const expectedWasmBindgen = `wasm-bindgen ${wasmBindgenVersion}`;
  if (installedWasmBindgen !== expectedWasmBindgen) {
    panic(`Expected ${expectedWasmBindgen}; found ${installedWasmBindgen}`);
  }
};

/**
 * The artifact loads from the package that ships it, so wasm-bindgen's bare
 * specifier has to become an explicitly relative one; a bundler resolves the
 * bare form against node_modules and finds nothing.
 */
const rewriteToPackageRelativeUrl = (artifact: RustWasmArtifact, directory: string): void => {
  const generatedJavaScriptPath = path.join(directory, `${artifact.outName}.js`);
  const generatedJavaScript = readFileSync(generatedJavaScriptPath, "utf8");
  const packageRelativeJavaScript = generatedJavaScript.replace(
    `new URL('${artifact.outName}_bg.wasm', import.meta.url)`,
    `new URL('./${artifact.outName}_bg.wasm', import.meta.url)`,
  );
  if (packageRelativeJavaScript === generatedJavaScript) {
    panic("Could not normalize wasm-bindgen's package-relative WebAssembly URL");
  }
  writeFileSync(
    generatedJavaScriptPath,
    `// @ts-nocheck -- generated and verified by ${path.basename(artifact.buildScript)}\n${packageRelativeJavaScript}`,
  );
};

const optimize = (artifact: RustWasmArtifact, wasmPath: string): void => {
  const wasmModule = binaryen.readBinary(readFileSync(wasmPath));
  binaryen.setOptimizeLevel(3);
  binaryen.setShrinkLevel(2);
  wasmModule.setFeatures(
    wasmModule.getFeatures() | binaryen.Features.BulkMemory | binaryen.Features.BulkMemoryOpt,
  );
  wasmModule.optimize();
  wasmModule.optimize();
  if (wasmModule.getMemoryInfo().shared) {
    panic(`${artifact.label} WebAssembly must not require shared memory or threads`);
  }
  writeFileSync(wasmPath, wasmModule.emitBinary());
  wasmModule.dispose();
};

type SizeReport = { readonly wasmBytes: number; readonly brotliBytes: number };

const assertLoadableWithinBudget = (artifact: RustWasmArtifact, wasmPath: string): SizeReport => {
  const wasmBytes = readFileSync(wasmPath);
  const brotliBytes = brotliCompressSync(wasmBytes, {
    params: { [zlibConstants.BROTLI_PARAM_QUALITY]: 9 },
  });
  if (!WebAssembly.validate(wasmBytes)) {
    panic(`Generated ${artifact.label} artifact is not valid WebAssembly`);
  }
  const compiledModule = new WebAssembly.Module(wasmBytes);
  const wasiImports = WebAssembly.Module.imports(compiledModule).filter(({ module }) =>
    module.startsWith("wasi"),
  );
  if (wasiImports.length > 0) {
    panic(
      `${artifact.label} WebAssembly must not import WASI: ${wasiImports.map(({ module, name }) => `${module}.${name}`).join(", ")}`,
    );
  }
  if (wasmBytes.byteLength > artifact.maximumWasmBytes) {
    panic(
      `${artifact.label} WebAssembly is ${wasmBytes.byteLength} bytes; budget is ${artifact.maximumWasmBytes}`,
    );
  }
  if (brotliBytes.byteLength > artifact.maximumBrotliBytes) {
    panic(
      `${artifact.label} Brotli size is ${brotliBytes.byteLength} bytes; budget is ${artifact.maximumBrotliBytes}`,
    );
  }
  return { wasmBytes: wasmBytes.byteLength, brotliBytes: brotliBytes.byteLength };
};

const assertCommittedMatches = (
  artifact: RustWasmArtifact,
  temporaryDir: string,
  generatedFiles: readonly string[],
  wasmFile: string,
): void => {
  const generatedDir = path.join(repoRoot, artifact.generatedDirectory);
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
      `Generated ${artifact.label} artifacts are stale: ${stale.join(", ")}. Run ${artifact.regenerateCommand}.`,
    );
  }
  if (!canonicalArtifactPlatform) {
    const committedWasmPath = path.join(generatedDir, wasmFile);
    if (!existsSync(committedWasmPath) || !WebAssembly.validate(readFileSync(committedWasmPath))) {
      panic(`Committed ${artifact.label} artifact is not valid WebAssembly`);
    }
  }
};

/**
 * Builds the crate for `wasm32-unknown-unknown`, optimizes it, and either
 * writes the result into the package or verifies that what is committed is what
 * this build produces.
 */
export const buildRustWasmArtifact = async (
  artifact: RustWasmArtifact,
  mode: BuildMode,
): Promise<void> => {
  const generatedDir = path.join(repoRoot, artifact.generatedDirectory);
  const temporaryDir = mkdtempSync(path.join(tmpdir(), `folio-${artifact.outName}-wasm-`));
  const cargoOutput = path.join(
    repoRoot,
    "target",
    "wasm32-unknown-unknown",
    "release",
    `${artifact.cargoArtifact}.wasm`,
  );
  const wasmFile = `${artifact.outName}_bg.wasm`;
  const generatedFiles = [
    `${artifact.outName}.js`,
    `${artifact.outName}.d.ts`,
    wasmFile,
    `${wasmFile}.d.ts`,
    `${artifact.outName}.inputs.sha256`,
  ] as const;

  await assertPinnedWasmBindgen();

  try {
    await $`cargo build --locked --release --target wasm32-unknown-unknown --features wasm -p ${artifact.crate}`.cwd(
      repoRoot,
    );
    await $`wasm-bindgen ${cargoOutput} --target web --remove-name-section --remove-producers-section --out-dir ${temporaryDir} --out-name ${artifact.outName}`.cwd(
      repoRoot,
    );

    rewriteToPackageRelativeUrl(artifact, temporaryDir);
    const wasmPath = path.join(temporaryDir, wasmFile);
    optimize(artifact, wasmPath);
    writeFileSync(
      path.join(temporaryDir, `${artifact.outName}.inputs.sha256`),
      generationInputDigest(artifact),
    );

    const { wasmBytes, brotliBytes } = assertLoadableWithinBudget(artifact, wasmPath);

    if (mode === "--write") {
      mkdirSync(generatedDir, { recursive: true });
      for (const file of generatedFiles) {
        writeFileSync(path.join(generatedDir, file), readFileSync(path.join(temporaryDir, file)));
      }
    } else {
      assertCommittedMatches(artifact, temporaryDir, generatedFiles, wasmFile);
    }

    const generatedFileSet: ReadonlySet<string> = new Set(generatedFiles);
    const unexpected = readdirSync(temporaryDir).filter((file) => !generatedFileSet.has(file));
    if (unexpected.length > 0) {
      panic(`Unexpected wasm-bindgen artifacts: ${unexpected.join(", ")}`);
    }
    console.log(
      `${artifact.label} WebAssembly: ${wasmBytes} bytes raw, ${brotliBytes} bytes Brotli`,
    );
  } finally {
    rmSync(temporaryDir, { recursive: true, force: true });
  }
};
