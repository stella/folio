// Deliberately violating fixture for scripts/no-untranslated-jsx-literal.test.ts.
// The raw JSX text child below must raise exactly one
// no-untranslated-jsx-literal diagnostic when linted with the repo config.
export const UntranslatedFixture = () => <button type="button">Save document</button>;
