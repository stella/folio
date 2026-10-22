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

type CargoMetadata = {
  readonly packages: readonly {
    readonly id: string;
    readonly name: string;
    readonly version: string;
    readonly source: string | null;
    readonly checksum: string | null;
  }[];
  readonly resolve: {
    readonly nodes: readonly {
      readonly id: string;
      readonly deps: readonly {
        readonly pkg: string;
        readonly dep_kinds: readonly {
          readonly kind: "dev" | "build" | null;
          readonly target?: string | null;
        }[];
      }[];
    }[];
  } | null;
};

/** Returns the canonical locked packages used by a crate's WebAssembly build. */
export const resolvedWasmBuildGraphDescription = (
  metadata: CargoMetadata,
  crate: string,
): string => {
  const roots = metadata.packages.filter(({ name }) => name === crate);
  if (roots.length !== 1) {
    panic(`Expected exactly one Cargo package named ${crate}; found ${roots.length}`);
  }
  const root = roots.at(0) ?? panic(`Cargo metadata does not contain package ${crate}`);
  if (metadata.resolve === null) {
    panic("Cargo metadata does not contain a dependency graph");
  }

  const nodes = new Map(metadata.resolve.nodes.map((node) => [node.id, node]));
  const packages = new Map(metadata.packages.map((package_) => [package_.id, package_]));
  const pending = [root.id];
  const visited = new Set<string>();
  while (pending.length > 0) {
    const id = pending.pop() ?? panic("Dependency traversal unexpectedly ran out of packages");
    if (visited.has(id)) {
      continue;
    }
    visited.add(id);
    const node = nodes.get(id) ?? panic(`Cargo metadata does not contain resolve node ${id}`);
    for (const dependency of node.deps) {
      if (dependency.dep_kinds.some(({ kind }) => kind === null || kind === "build")) {
        pending.push(dependency.pkg);
      }
    }
  }

  const resolved = [...visited].map((id) => {
    const package_ = packages.get(id) ?? panic(`Cargo metadata does not contain package ${id}`);
    return {
      name: package_.name,
      version: package_.version,
      source: package_.source,
      checksum: package_.checksum,
    };
  });
  resolved.sort((left, right) => {
    const leftDescription = JSON.stringify(left);
    const rightDescription = JSON.stringify(right);
    return leftDescription < rightDescription ? -1 : Number(leftDescription > rightDescription);
  });
  return `${JSON.stringify(resolved)}\n`;
};

const canonicalJson = (value: unknown): string => {
  if (Array.isArray(value)) {
    return `[${value.map(canonicalJson).join(",")}]`;
  }
  if (typeof value === "object" && value !== null) {
    const entries = Object.entries(value).toSorted(([left], [right]) =>
      left < right ? -1 : Number(left > right),
    );
    return `{${entries.map(([key, entry]) => `${JSON.stringify(key)}:${canonicalJson(entry)}`).join(",")}}`;
  }
  return JSON.stringify(value);
};

/**
 * Returns a crate manifest as the WebAssembly build reads it: dev-dependencies,
 * including target-specific ones, never reach the build.
 */
export const manifestBuildDescription = (manifest: string): string => {
  const parsed: unknown = Bun.TOML.parse(manifest);
  if (typeof parsed !== "object" || parsed === null || Array.isArray(parsed)) {
    panic("Cargo manifest must be a table");
  }
  const {
    "dev-dependencies": _devDependencies,
    target,
    ...build
  } = parsed as Record<string, unknown>;
  if (typeof target === "object" && target !== null && !Array.isArray(target)) {
    build["target"] = Object.fromEntries(
      Object.entries(target).map(([platform, table]) => {
        if (typeof table !== "object" || table === null || Array.isArray(table)) {
          return [platform, table];
        }
        const { "dev-dependencies": _platformDevDependencies, ...platformBuild } = table as Record<
          string,
          unknown
        >;
        return [platform, platformBuild];
      }),
    );
  } else if (target !== undefined) {
    build["target"] = target;
  }
  return `${canonicalJson(build)}\n`;
};

const generationInputsOf = (artifact: RustWasmArtifact): string[] => {
  const sourceDirectory = path.join(repoRoot, artifact.crateDirectory, "src");
  const sourceFiles = readdirSync(sourceDirectory, { recursive: true })
    .filter((file): file is string => typeof file === "string" && file.endsWith(".rs"))
    .map((file) => path.join(sourceDirectory, file));
  return [
    path.join(repoRoot, "Cargo.toml"),
    path.join(repoRoot, "rust-toolchain.toml"),
    path.join(repoRoot, "scripts", "lib", "rust-wasm-artifact.ts"),
    path.join(repoRoot, artifact.buildScript),
    ...sourceFiles,
  ].sort();
};

const generationInputDigest = async (artifact: RustWasmArtifact): Promise<string> => {
  const digest = createHash("sha256");
  for (const file of generationInputsOf(artifact)) {
    digest.update(path.relative(repoRoot, file));
    digest.update("\0");
    digest.update(readFileSync(file));
    digest.update("\0");
  }
  digest.update(path.join(artifact.crateDirectory, "Cargo.toml"));
  digest.update("\0");
  digest.update(
    manifestBuildDescription(
      readFileSync(path.join(repoRoot, artifact.crateDirectory, "Cargo.toml"), "utf8"),
    ),
  );
  digest.update("\0");
  // Unfiltered by platform: build scripts and procedural macros run on the
  // build host, so a host-only dependency of theirs is a build input too.
  const metadata = JSON.parse(
    await capture("cargo", ["metadata", "--locked", "--format-version", "1", "--features", "wasm"]),
  ) as CargoMetadata;
  digest.update("resolved-cargo-build-graph");
  digest.update("\0");
  digest.update(resolvedWasmBuildGraphDescription(metadata, artifact.crate));
  digest.update("\0");
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
      await generationInputDigest(artifact),
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
