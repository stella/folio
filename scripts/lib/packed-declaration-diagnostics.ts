import { readFileSync } from "node:fs";
import { builtinModules } from "node:module";
import path from "node:path";
import ts from "typescript";

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

const referencesBuiltin = (diagnostic: ts.Diagnostic) => {
  if (
    !diagnostic.file ||
    diagnostic.start === undefined ||
    diagnostic.length === undefined
  )
    return false;
  const specifier = diagnostic.file.text
    .slice(diagnostic.start, diagnostic.start + diagnostic.length)
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
export const declarationResolutionDiagnostics = ({
  projectFile,
  packageNames,
}: DeclarationResolutionOptions) => {
  const project = ts.parseConfigFileTextToJson(
    projectFile,
    readFileSync(projectFile, "utf8"),
  );
  if (project.error) return [project.error];
  const config = ts.parseJsonConfigFileContent(
    project.config,
    ts.sys,
    path.dirname(projectFile),
  );
  if (config.errors.length > 0) return config.errors;
  const program = ts.createProgram({
    rootNames: config.fileNames,
    options: { ...config.options, skipLibCheck: false },
  });
  return program.getSourceFiles().flatMap((file) => {
    if (!file.isDeclarationFile) return [];
    const filename = file.fileName.replaceAll("\\", "/");
    if (
      !packageNames.some((name) => filename.includes(`/node_modules/${name}/`))
    )
      return [];
    return program
      .getSemanticDiagnostics(file)
      .filter(
        (diagnostic) =>
          RESOLUTION_ERRORS.has(diagnostic.code) &&
          !referencesBuiltin(diagnostic),
      );
  });
};

if (import.meta.main) {
  const { PUBLISHED_PACKAGES } = await import("./published-packages");
  const projectFile = process.argv.at(2);
  if (!projectFile)
    throw new TypeError(
      "usage: packed-declaration-diagnostics.ts <tsconfig.json>",
    );
  const diagnostics = declarationResolutionDiagnostics({
    projectFile,
    packageNames: PUBLISHED_PACKAGES.map(({ name }) => name),
  });
  if (diagnostics.length > 0) {
    process.stderr.write(
      ts.formatDiagnostics(diagnostics, {
        getCanonicalFileName: (file) => file,
        getCurrentDirectory: () => process.cwd(),
        getNewLine: () => "\n",
      }),
    );
    process.exitCode = 1;
  }
}
