import type { TsdownPlugin } from "tsdown";

export const STABLE_RELATIVE_IMPORT_ORDER_PLUGIN = "stable-relative-import-order";

type RelativeImport = {
  start: number;
  end: number;
  source: string;
  group: number;
};

const compareCodeUnits = (left: string, right: string): number => {
  if (left < right) {
    return -1;
  }
  if (left > right) {
    return 1;
  }
  return 0;
};

/**
 * Canonicalize only consecutive relative-import groups. External imports and
 * their ordering boundaries stay untouched; callers must additionally limit
 * this to source-mirrored packages that declare `sideEffects: false`.
 */
export const canonicalizeRelativeImportOrder = (
  code: string,
  imports: readonly RelativeImport[],
): string => {
  if (imports.length < 2) {
    return code;
  }

  const sortedByGroup = new Map<number, { code: string; source: string }[]>();
  for (const importDeclaration of imports) {
    const group = sortedByGroup.get(importDeclaration.group) ?? [];
    group.push({
      code: code.slice(importDeclaration.start, importDeclaration.end),
      source: importDeclaration.source,
    });
    sortedByGroup.set(importDeclaration.group, group);
  }
  for (const group of sortedByGroup.values()) {
    group.sort(
      (left, right) =>
        compareCodeUnits(left.source, right.source) || compareCodeUnits(left.code, right.code),
    );
  }

  let cursor = 0;
  let result = "";
  const groupIndexes = new Map<number, number>();
  for (const statement of imports) {
    const groupIndex = groupIndexes.get(statement.group) ?? 0;
    const replacement = sortedByGroup.get(statement.group)?.at(groupIndex);
    if (!replacement) {
      return code;
    }
    groupIndexes.set(statement.group, groupIndex + 1);
    result += code.slice(cursor, statement.start);
    result += replacement.code;
    cursor = statement.end;
  }
  return result + code.slice(cursor);
};

export const stableRelativeImportOrder = (): TsdownPlugin => ({
  name: STABLE_RELATIVE_IMPORT_ORDER_PLUGIN,
  renderChunk(code, _chunk, options) {
    if (options.sourcemap) {
      this.error(`${STABLE_RELATIVE_IMPORT_ORDER_PLUGIN} requires sourcemaps to remain disabled`);
    }

    const imports: RelativeImport[] = [];
    let group = 0;
    let previousWasRelative = false;
    for (const statement of this.parse(code).body) {
      if (statement.type === "ImportDeclaration" && statement.source.value.startsWith(".")) {
        if (!previousWasRelative) {
          group += 1;
        }
        imports.push({
          start: statement.start,
          end: statement.end,
          source: statement.source.value,
          group,
        });
        previousWasRelative = true;
        continue;
      }
      previousWasRelative = false;
    }

    const canonical = canonicalizeRelativeImportOrder(code, imports);
    if (canonical === code) {
      return null;
    }
    return { code: canonical, map: null };
  },
});
