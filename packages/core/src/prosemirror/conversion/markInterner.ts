import { panic } from "better-result";
import type { Attrs, Mark, MarkType, Schema } from "prosemirror-model";

import type { MarkFactory } from "../extensions/marks/markUtils";

/**
 * Leaves ProseMirror compares with `===`, minus `-0`/`0`: sharing a mark makes
 * `Mark.eq` take its identity shortcut, so two leaves may share only when
 * ProseMirror would call them equal anyway. `NaN` never equals itself there.
 */
const sameLeaf = (left: unknown, right: unknown): boolean =>
  left === right && Object.is(left, right);

/**
 * Whether two attr values are interchangeable: `sameLeaf` on leaves, and the
 * same own keys in the same order, so `-0`/`0`, `NaN` and `undefined`/absent
 * stay apart. Output follows the shared value's key order, which is why order
 * counts.
 */
const sameAttrValue = (left: unknown, right: unknown): boolean => {
  if (sameLeaf(left, right)) {
    return true;
  }
  if (typeof left !== "object" || typeof right !== "object" || left === null || right === null) {
    return false;
  }
  if (Object.getPrototypeOf(left) !== Object.getPrototypeOf(right)) {
    return false;
  }
  if (Array.isArray(left)) {
    if (!Array.isArray(right) || left.length !== right.length) {
      return false;
    }
    for (let index = 0; index < left.length; index += 1) {
      if (index in left !== index in right || !sameAttrValue(left[index], right[index])) {
        return false;
      }
    }
    return true;
  }
  const leftKeys = Object.keys(left);
  const rightKeys = Object.keys(right);
  return (
    leftKeys.length === rightKeys.length &&
    leftKeys.every(
      (key, index) =>
        key === rightKeys[index] && sameAttrValue(Reflect.get(left, key), Reflect.get(right, key)),
    )
  );
};

type InternedMark = {
  mark: Mark;
  /** Declared attrs whose value is not the attr's default. */
  nonDefault: ReadonlySet<string>;
};

const internedMark = (type: MarkType, attrs: Attrs | null | undefined): InternedMark => {
  const mark = type.create(attrs);
  const nonDefault = new Set<string>();
  for (const [name, spec] of Object.entries(type.spec.attrs ?? {})) {
    if (!sameLeaf(mark.attrs[name], spec.default)) {
      nonDefault.add(name);
    }
  }
  return { mark, nonDefault };
};

/**
 * Whether `create(attrs)` would compute attrs interchangeable with the
 * interned mark's. `create` fills every declared attr `attrs` leaves
 * `undefined` from its default, so the given attrs must match each value they
 * set and set every attr where the mark holds something else.
 */
const matchesInterned = (
  type: MarkType,
  attrs: Attrs | null | undefined,
  { mark, nonDefault }: InternedMark,
): boolean => {
  const specs = type.spec.attrs ?? {};
  let covered = 0;
  for (const name in attrs) {
    const given: unknown = attrs[name];
    if (given === undefined || !Object.hasOwn(specs, name)) {
      continue;
    }
    if (!sameAttrValue(given, mark.attrs[name])) {
      return false;
    }
    if (nonDefault.has(name)) {
      covered += 1;
    }
  }
  return covered === nonDefault.size;
};

/**
 * Build marks that share one instance per distinct (type, attrs) value.
 *
 * A document repeats a small set of run formattings across thousands of runs,
 * so one conversion builds each distinct mark once. Marks are immutable, and
 * ProseMirror compares them by value, so sharing an instance changes nothing a
 * caller can observe except the saved allocation. Create one interner per
 * conversion; it holds every mark it built until it is dropped.
 *
 * The native JSON serialization is only a lookup key: it collapses values the
 * mark would keep apart (`-0`, `NaN`, nested `undefined`), so a hit is shared
 * only after an exact comparison, and a mismatch builds an unshared mark. A
 * mark holding `NaN` is never shared: ProseMirror does not find two of them
 * equal, and a shared instance would.
 */
export const createMarkInterner = (schema: Schema): MarkFactory => {
  const marksByType = new Map<MarkType, Map<string, InternedMark>>();
  return (typeName, attrs) => {
    const type = schema.marks[typeName];
    if (type === undefined) {
      panic(`Unknown mark type ${typeName}.`);
    }
    let marks = marksByType.get(type);
    if (marks === undefined) {
      marks = new Map();
      marksByType.set(type, marks);
    }
    const key = JSON.stringify(attrs ?? null);
    const cached = marks.get(key);
    if (cached === undefined) {
      const interned = internedMark(type, attrs);
      marks.set(key, interned);
      return interned.mark;
    }
    return matchesInterned(type, attrs, cached) ? cached.mark : type.create(attrs);
  };
};
