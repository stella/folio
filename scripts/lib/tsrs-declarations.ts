import { createRequire } from "node:module";
import path from "node:path";

export type TsrsDeclarations = {
  readonly generator: "tsgo";
  readonly tsconfig: string;
  readonly tsgo: { readonly path: string };
};

/**
 * Declaration options for a package whose published declarations carry doc
 * comments the oxc generator drops: a comment on a union member, or a
 * detached comment above an export. tsrs (TypeScript 7) emits them as the
 * compiler does. The generator is named explicitly, because tsdown left to
 * choose falls back to an installed `typescript` package.
 *
 * `tsconfig` names a config that includes only the package's published
 * sources: the compiler emits a declaration for every file in its program,
 * and one outside the package (a typecheck fixture under the repository's
 * `test/`) would be written beside its source.
 */
export const tsrsDeclarations = (tsconfig: string): TsrsDeclarations => ({
  generator: "tsgo",
  tsconfig,
  tsgo: {
    path: path.join(
      path.dirname(createRequire(import.meta.url).resolve("tsrs/package.json")),
      "bin",
      "tsrs",
    ),
  },
});
