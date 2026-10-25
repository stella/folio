import { readFileSync } from "node:fs";
import { builtinModules } from "node:module";
import path from "node:path";
import {
  API,
  type Diagnostic,
  formatDiagnostics,
  type FormatDiagnosticsHost,
} from "tsrs/unstable/async";

// Module/type resolution diagnostics, including named/default export mismatches.
// Other semantic errors in third-party types are outside this dependency gate.
const RESOLUTION_ERRORS = new Set([
  2305, // Module has no exported member.
  2306, // Referenced file is not a module.
  2307, // Module or its declarations cannot be found.
  2459,
  2460, // Referenced member is declared locally but not exported under that name.
  2613,
  2614,
  2724, // Missing default/named exports, including suggested alternatives.
  2688, // Type reference directive cannot be resolved.
  2694, // Imported namespace has no referenced member.
  2792, // Module resolution failed with a suggested resolution mode.
  7016, // Resolved JavaScript module has no declarations.
]);

const BUILTINS = new Set(
  builtinModules.flatMap((name) => [name, `node:${name}`]),
);

const DECLARATION_FILE = /\.d\.(?:[cm]?ts|[^.]+\.ts)$/u;

const referencesBuiltin = (diagnostic: Diagnostic) => {
  if (diagnostic.fileName === undefined) return false;
  const specifier = readFileSync(diagnostic.fileName, "utf8")
    .slice(diagnostic.pos, diagnostic.end)
    .replace(/^["']|["']$/gu, "");
  // Match the manifest scanner's explicit host-builtin exemption. Host ambient
  // Node declarations are not runtime dependencies of the portable packages.
  return BUILTINS.has(specifier);
};

type DeclarationResolutionOptions = {
  projectFile: string;
  packageNames: readonly string[];
};

/** Check our installed declarations with library checking enabled and explicit ownership. */
export const declarationResolutionDiagnostics = async ({
  projectFile,
  packageNames,
}: DeclarationResolutionOptions): Promise<readonly Diagnostic[]> => {
  const api = new API({ cwd: path.dirname(projectFile) });
  try {
    const config = await api.parseConfigFile(projectFile);
    if (config.errors.length > 0) return config.errors;
    const snapshot = await api.createSnapshot({
      createPrograms: [
        {
          rootFiles: config.fileNames,
          compilerOptions: { ...config.options, skipLibCheck: false },
        },
      ],
    });
    const [program] = snapshot.operation.createdPrograms;
    const owned = (await program.getSourceFileNames()).filter((fileName) => {
      const filename = fileName.replaceAll("\\", "/");
      return (
        DECLARATION_FILE.test(filename) &&
        packageNames.some((name) => filename.includes(`/node_modules/${name}/`))
      );
    });
    if (owned.length === 0) return [];
    return (await program.getSemanticDiagnostics(owned)).filter(
      (diagnostic) =>
        RESOLUTION_ERRORS.has(diagnostic.code) &&
        !referencesBuiltin(diagnostic),
    );
  } finally {
    await api.close();
  }
};

const FORMAT_HOST: FormatDiagnosticsHost = {
  getCanonicalFileName: (file) => file,
  getCurrentDirectory: () => process.cwd(),
  getNewLine: () => "\n",
};

if (import.meta.main) {
  const { PUBLISHED_PACKAGES } = await import("./published-packages");
  const projectFile = process.argv.at(2);
  if (!projectFile)
    throw new TypeError(
      "usage: packed-declaration-diagnostics.ts <tsconfig.json>",
    );
  const diagnostics = await declarationResolutionDiagnostics({
    projectFile,
    packageNames: PUBLISHED_PACKAGES.map(({ name }) => name),
  });
  if (diagnostics.length > 0) {
    process.stderr.write(formatDiagnostics(diagnostics, FORMAT_HOST));
    process.exitCode = 1;
  }
}
