import base from "./oxlint.config";

// The reserved-value rule turned on over published source.
//
// `bun run lint` leaves it off: the repository carries a few hundred bare
// comparisons predating the registry, and turning them all into errors at once
// would only force a blanket suppression. `scripts/reserved-value-baseline.ts`
// lints with this config instead and holds the count per file to a committed
// baseline that may only shrink, the way the React Compiler bailout guard does.
export default {
  ...base,
  overrides: [
    ...(base.overrides ?? []),
    {
      files: ["packages/*/src/**/*.{ts,tsx}"],
      rules: {
        "folio-reserved-values/no-bare-reserved-compare": "error",
      },
    },
    {
      // A test names a sentinel to assert on it; that is the value under test,
      // not a second reader.
      files: [
        "packages/*/src/**/*.test.{ts,tsx}",
        "packages/*/src/**/*.property.test.{ts,tsx}",
        "packages/*/src/**/__tests__/**/*.{ts,tsx}",
        "packages/*/src/**/__fixtures__/**/*.{ts,tsx}",
      ],
      rules: {
        "folio-reserved-values/no-bare-reserved-compare": "off",
      },
    },
  ],
};
