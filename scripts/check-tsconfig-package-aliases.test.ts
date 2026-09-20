import { describe, expect, test } from "bun:test";

import { validateTSConfigPackageAliases } from "./check-tsconfig-package-aliases";

const PACKAGES = [
  { directory: "core", name: "@stll/folio-core" },
  { directory: "react", name: "@stll/folio-react" },
];

describe("validateTSConfigPackageAliases", () => {
  test("accepts a wildcard and a specific-subpath alias that target the package's own src", () => {
    const issues = validateTSConfigPackageAliases(PACKAGES, {
      "@stll/folio-core": ["packages/core/src/index.ts"],
      "@stll/folio-core/*": ["packages/core/src/*"],
    });
    expect(issues).toEqual([]);
  });

  test("rejects an alias that isn't a @stll/<package> specifier", () => {
    const issues = validateTSConfigPackageAliases(PACKAGES, {
      "@scratch-non-package-alias/*": ["packages/core/src/*"],
    });
    expect(issues).toHaveLength(1);
    expect(issues[0]).toContain('"@scratch-non-package-alias/*"');
    expect(issues[0]).toContain("is not a @stll/<package> alias");
  });

  test("rejects an alias for a package that doesn't exist", () => {
    const issues = validateTSConfigPackageAliases(PACKAGES, {
      "@stll/does-not-exist/*": ["packages/core/src/*"],
    });
    expect(issues).toHaveLength(1);
    expect(issues[0]).toContain("references unknown package");
  });

  test("rejects an alias whose target escapes the package's own src", () => {
    const issues = validateTSConfigPackageAliases(PACKAGES, {
      "@stll/folio-core/*": ["packages/react/src/*"],
    });
    expect(issues).toHaveLength(1);
    expect(issues[0]).toContain("must stay under packages/core/src/");
  });
});
