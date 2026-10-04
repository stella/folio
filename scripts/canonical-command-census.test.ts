import { expect, test } from "bun:test";
import fc from "fast-check";
import { assertProperty } from "../test/property-testing";
import {
  censusUndescribedCanonicalCommands,
  checkCanonicalCommandBaseline,
} from "./lib/canonical-command-census";

const owner = "packages/core/src/prosemirror/canonicalCommands.ts";
const registry = "packages/core/src/prosemirror/extensions/core/Fixture.ts";
const census = (source: string, helper = "") =>
  censusUndescribedCanonicalCommands([
    { file: owner, source: "export const withCanonicalCommand = () => {};" },
    { file: registry, source },
    { file: "packages/core/src/prosemirror/helpers.ts", source: helper },
  ]);
const imports = 'import { withCanonicalCommand as wrap } from "../../canonicalCommands";';

test("descriptor proof follows factory returns and the canonical wrapper owner", () => {
  const cases = {
    wrapped: "() => wrap(raw)",
    returnedLocal: "() => { const command = wrap(raw); return command; }",
    wrappedBranches: "() => flag ? wrap(raw) : wrap(other)",
    importedHelper: "() => described()",
    raw: "() => raw",
    incidental: "() => { wrap(raw); return raw; }",
    nested: "() => { const nested = () => wrap(raw); return raw; }",
    mixedBranches: "() => { if (flag) return wrap(raw); return raw; }",
    recursive: "() => cycle()",
    ambiguous: "() => duplicate()",
  };
  const registrations = Object.entries(cases)
    .map(([name, factory]) => `${name}: ${factory}`)
    .join(",");
  const source = `${imports}
    import { described } from "../../helpers";
    const cycle = () => cycle();
    const duplicate = () => wrap(raw);
    function duplicate() { return raw; }
    const extension = { commands: { ${registrations} } };`;
  expect(
    census(
      source,
      `${imports.replace("../../canonicalCommands", "./canonicalCommands")} export const described = () => wrap(raw);`,
    ),
  ).toEqual(
    ["raw", "incidental", "nested", "mixedBranches", "recursive", "ambiguous"]
      .map((name) => `${registry}#${name}`)
      .sort(),
  );
  const untrustedImports = [
    'import { withCanonicalCommand as wrap } from "unrelated";',
    'import { unrelated as wrap } from "../../canonicalCommands";',
    "const wrap = () => raw;",
  ];
  for (const statement of untrustedImports) {
    expect(
      census(`${statement} const extension = { commands: { wrapped: () => wrap(raw) } };`),
    ).toEqual([`${registry}#wrapped`]);
  }
});

test("registry additions and descriptor removals can only enlarge the derived remaining set", () => {
  assertProperty(
    fc.property(
      fc.uniqueArray(fc.integer({ min: 0, max: 1000 }), { minLength: 1, maxLength: 12 }),
      (ids) => {
        const names = ids.map((id) => `command${id}`);
        const registryFor = (wrapped: boolean) =>
          `${imports} const extension = { commands: { ${names.map((name) => `${name}: () => ${wrapped ? "wrap(raw)" : "raw"}`).join(",")} } };`;
        const remaining = census(registryFor(false));
        expect(census(registryFor(true))).toEqual([]);
        expect(remaining).toEqual(names.map((name) => `${registry}#${name}`).sort());
        expect(checkCanonicalCommandBaseline([], remaining)).toEqual([]);
        expect(checkCanonicalCommandBaseline(remaining, [])).toEqual(remaining);
        expect(checkCanonicalCommandBaseline(remaining.concat(remaining), remaining)).toEqual(
          remaining,
        );
      },
    ),
    { numRuns: 25 },
  );
});

test("registry shapes cannot hide uncensused registrations", () => {
  expect(() => census("const extension = { commands }; ")).toThrow("census");
  expect(() => census("const extension = { commands() { return other; } }; ")).toThrow("census");
  for (const commands of ["other", "{ ...other }", "{ [dynamic]: () => raw }"]) {
    expect(() => census(`const extension = { commands: ${commands} };`)).toThrow("census");
  }
});
