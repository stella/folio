/**
 * One writer for every `CT_OnOff` element in the tree.
 *
 * A toggle has three states: absent, an explicit on, an explicit off. Written
 * by hand, it reliably loses one of them — `if (value) parts.push("<w:x/>")`
 * writes the on and drops the off, and an explicit off is the only thing that
 * cancels the on a style above it turns on. Seven serializers had that shape at
 * once, so the shape is the defect, not any of the seven.
 *
 * The rule is therefore structural: `serializeOnOffElement` in
 * `packages/docx-core/src/serialize/xml.ts` is the only place in the tree that
 * may spell a `CT_OnOff` element, and everything else calls it. This test reads
 * the element names from the committed schema graph rather than from a list, so
 * a toggle the format adds is covered the day the graph learns about it.
 *
 * It scans emitted markup, not prose: a parser's comment about `w:val="off"` is
 * documentation of the lexical space, not a second writer.
 */

import { describe, expect, test } from "bun:test";
import path from "node:path";

const REPOSITORY_ROOT = path.resolve(import.meta.dir, "..");
const PACKAGES_DIR = path.join(REPOSITORY_ROOT, "packages");
const SCHEMA_PATH = path.join(
  REPOSITORY_ROOT,
  "specifications/generated/docx-transitional-schema.gen.json",
);

/**
 * The one module allowed to spell a `CT_OnOff` element.
 *
 * Nothing else belongs here. A second entry would be a second writer, which is
 * the thing this test exists to refuse.
 */
const WRITER = "docx-core/src/serialize/xml.ts";

type SchemaGraph = { children: ReadonlyArray<{ name?: string; type?: string }> };

const onOffElementNames = async (): Promise<ReadonlySet<string>> => {
  const graph = (await Bun.file(SCHEMA_PATH).json()) as SchemaGraph;
  const names = new Set<string>();
  for (const child of graph.children) {
    if (child.name !== undefined && child.type?.includes("CT_OnOff") === true) {
      names.add(child.name);
    }
  }
  return names;
};

/**
 * A literal `w:` element emitted either bare or with a `w:val`.
 *
 * Anchored on `<w:` plus the element name so a field named after the element
 * (`formatting.hidden`) and a sentence mentioning one do not match; only markup
 * a serializer hands to a consumer does.
 */
const EMITTED_ELEMENT = /<w:(?<name>[A-Za-z0-9]+)\s*(?:\/>|w:val=)/gu;

/** A line that is only a comment cannot be a writer. */
const isComment = (line: string): boolean => /^\s*(?:\/\/|\/\*|\*)/u.test(line);

/**
 * A generic toggle writer: an element whose name is interpolated, spelled with
 * the explicit-off `w:val`.
 *
 * {@link EMITTED_ELEMENT} cannot see this shape, because there is no element
 * name to check against the schema. It is the shape three modules had before
 * the writer existed, each a copy of the other two, so it is refused by
 * spelling rather than by name.
 */
const INTERPOLATED_ON_OFF = /<w:\$\{[^}]+\}\s+w:val="0"/u;

const isProductSource = (relative: string): boolean =>
  !relative.includes("/__tests__/") &&
  !relative.includes("/__fixtures__/") &&
  !relative.endsWith(".test.ts") &&
  !relative.endsWith(".test.tsx");

const handBuiltOnOffElements = async (): Promise<string[]> => {
  const names = await onOffElementNames();
  const offenders: string[] = [];
  for await (const file of new Bun.Glob("*/src/**/*.ts").scan({
    absolute: true,
    cwd: PACKAGES_DIR,
  })) {
    const relative = path.relative(PACKAGES_DIR, file);
    if (!isProductSource(`/${relative}`) || relative === WRITER) {
      continue;
    }
    const source = await Bun.file(file).text();
    for (const [index, line] of source.split("\n").entries()) {
      if (isComment(line)) {
        continue;
      }
      if (INTERPOLATED_ON_OFF.test(line)) {
        offenders.push(`${relative}:${index + 1}  a second generic toggle writer`);
      }
      for (const match of line.matchAll(EMITTED_ELEMENT)) {
        const name = match.groups?.["name"];
        if (name !== undefined && names.has(name)) {
          offenders.push(`${relative}:${index + 1}  <w:${name}>`);
        }
      }
    }
  }
  return offenders.sort();
};

describe("CT_OnOff element writer", () => {
  test("the schema graph names the toggles the scan looks for", async () => {
    const names = await onOffElementNames();
    // Anti-vacuity: a graph read that returned nothing would pass the scan
    // below for the wrong reason. These four are the toggles the seven broken
    // sites wrote, one per property set.
    expect(names.size).toBeGreaterThan(100);
    for (const name of ["cantSplit", "tblHeader", "noWrap", "specVanish"]) {
      expect(names.has(name)).toBe(true);
    }
  });

  test("the scan sees the shape it refuses", () => {
    const seen = (line: string): string[] =>
      isComment(line)
        ? []
        : [...line.matchAll(EMITTED_ELEMENT)].map((m) => m.groups?.["name"] ?? "");
    expect(seen('  parts.push("<w:cantSplit/>");')).toEqual(["cantSplit"]);
    expect(seen("  parts.push('<w:hideMark w:val=\"0\"/>');")).toEqual(["hideMark"]);
    // Prose about the markup is not markup.
    expect(seen(" * A row that carries `<w:cantSplit/>` may not break.")).toEqual([]);
    // A call through the writer spells no element at all.
    expect(seen('  pushOnOffElement(parts, formatting.cantSplit, "cantSplit");')).toEqual([]);
    // A generic writer has no element name to check, so it is refused by shape.
    expect(INTERPOLATED_ON_OFF.test('  return `<w:${name} w:val="0"/>`;')).toBe(true);
    expect(INTERPOLATED_ON_OFF.test('  return `<w:${name} w:type="dxa"/>`;')).toBe(false);
  });

  test("no serializer builds a CT_OnOff element by hand", async () => {
    expect(await handBuiltOnOffElements()).toEqual([]);
  });
});
