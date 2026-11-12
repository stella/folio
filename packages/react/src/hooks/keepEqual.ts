/**
 * State the editor host writes while the user types: caret and selection
 * geometry, overlay rectangles, toolbar formatting. Most writes carry a value
 * equal to the one already held (another empty rectangle list, the same
 * caret), and every React commit with the editor focused reads the editor's
 * whole DOM to save the selection. A write that changes nothing must not
 * commit, so these states keep the value they hold when the next one is
 * equal to it.
 */
import {
  type Dispatch,
  type SetStateAction,
  useCallback,
  useState,
} from "react";

const isPlainObject = (value: object): boolean => {
  const prototype: unknown = Object.getPrototypeOf(value);
  return prototype === Object.prototype || prototype === null;
};

/**
 * Whether two values hold the same data: primitives by `Object.is`, arrays,
 * maps, sets and plain objects by their contents. Other objects (DOM nodes,
 * class instances, functions) only when they are the same object.
 */
export const structurallyEqual = (left: unknown, right: unknown): boolean => {
  if (Object.is(left, right)) return true;
  if (
    typeof left !== "object" ||
    typeof right !== "object" ||
    left === null ||
    right === null
  )
    return false;
  if (Array.isArray(left) || Array.isArray(right)) {
    if (!Array.isArray(left) || !Array.isArray(right)) return false;
    return (
      left.length === right.length &&
      left.every((item, index) => structurallyEqual(item, right[index]))
    );
  }
  if (left instanceof Map || right instanceof Map) {
    if (!(left instanceof Map) || !(right instanceof Map)) return false;
    if (left.size !== right.size) return false;
    for (const [key, value] of left) {
      if (!right.has(key) || !structurallyEqual(value, right.get(key)))
        return false;
    }
    return true;
  }
  if (left instanceof Set || right instanceof Set) {
    if (!(left instanceof Set) || !(right instanceof Set)) return false;
    if (left.size !== right.size) return false;
    for (const value of left) if (!right.has(value)) return false;
    return true;
  }
  if (!isPlainObject(left) || !isPlainObject(right)) return false;
  const leftKeys = Object.keys(left);
  const rightKeys = Object.keys(right);
  return (
    leftKeys.length === rightKeys.length &&
    leftKeys.every(
      (key) =>
        Object.hasOwn(right, key) &&
        structurallyEqual(Reflect.get(left, key), Reflect.get(right, key)),
    )
  );
};

/**
 * The next value of a state that keeps its current value when the next one
 * is equal to it,
 * so an unchanged write never re-renders or commits.
 */
export function keepEqual<T>(previous: T, next: SetStateAction<T>): T {
  const resolved =
    typeof next === "function" ? (next as (previous: T) => T)(previous) : next;
  return structurallyEqual(previous, resolved) ? previous : resolved;
}

/**
 * `useState` for state the editor host writes while the user types: the
 * setter keeps the current value when the next one is equal to it. The
 * update runs through React's own state queue, which drops an unchanged
 * value before it schedules a render.
 */
export function useEqualState<T>(
  initial: T | (() => T),
): [T, Dispatch<SetStateAction<T>>] {
  const [value, setValue] = useState(initial);
  const setIfChanged = useCallback<Dispatch<SetStateAction<T>>>(
    (next) => setValue((previous) => keepEqual(previous, next)),
    [],
  );
  return [value, setIfChanged];
}
