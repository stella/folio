import base from "./oxlint.config";

// The identity-attribute rule turned on over published source.
//
// `bun run lint` leaves it off: the repository predates the derived identity
// set and still carries the prefix-resolved reads the rule rejects, so turning
// it on at once would only produce a file of suppressions.
// `scripts/identity-attribute-baseline.ts` lints with this config instead and
// holds the count per file to a committed baseline that may only shrink, the
// way the reserved-value guard does.
export default {
  ...base,
  overrides: [
    ...(base.overrides ?? []),
    {
      files: ["packages/*/src/**/*.{ts,tsx}"],
      rules: {
        "folio-identity-attributes/no-prefix-resolved-identity-read": "error",
      },
    },
    {
      // A test builds the element it reads back, so the prefix it names is the
      // one the fixture wrote; there is no foreign attribute to confuse it with.
      files: [
        "packages/*/src/**/*.test.{ts,tsx}",
        "packages/*/src/**/*.property.test.{ts,tsx}",
        "packages/*/src/**/__tests__/**/*.{ts,tsx}",
        "packages/*/src/**/__fixtures__/**/*.{ts,tsx}",
      ],
      rules: {
        "folio-identity-attributes/no-prefix-resolved-identity-read": "off",
      },
    },
  ],
};
