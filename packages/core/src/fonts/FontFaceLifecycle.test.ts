import { expect, mock, test } from "bun:test";

import { FontFaceLifecycle } from "./FontFaceLifecycle";

const completionOrders = [0, 1, 2].flatMap((first) =>
  [0, 1, 2]
    .filter((second) => first !== second)
    .map((second) => [first, second, 3 - first - second]),
);

for (const order of completionOrders) {
  test(`retains active fonts until the latest replacement is ready (${order.join(",")})`, () => {
    const initial = { family: "Initial Face" };
    const replacements = [
      { family: "Stale Face" },
      { family: "Other Stale Face" },
      { family: "Replacement Face" },
    ];
    const installed = new Set([initial]);
    const removals = new Map<(typeof replacements)[number], number>();
    const remove = (faces: readonly (typeof replacements)[number][]) => {
      for (const face of faces) {
        installed.delete(face);
        removals.set(face, (removals.get(face) ?? 0) + 1);
      }
    };
    const lifecycle = new FontFaceLifecycle(remove);
    const changed = mock(() => {});
    lifecycle.subscribe(changed);
    lifecycle.beginReplacement()([initial]);
    const pending = replacements.map((face) => ({
      face,
      commit: lifecycle.beginReplacement(),
    }));
    expect(installed).toEqual(new Set([initial]));
    expect(changed).toHaveBeenCalledTimes(1);

    let latestCompleted = false;
    for (const index of order) {
      const replacement = pending.at(index);
      if (!replacement) throw new Error("Missing replacement fixture");
      installed.add(replacement.face);
      replacement.commit([replacement.face]);
      latestCompleted ||= index === pending.length - 1;
      expect(installed).toEqual(
        new Set([latestCompleted ? replacements.at(-1) : initial]),
      );
      expect(changed).toHaveBeenCalledTimes(latestCompleted ? 2 : 1);
    }

    lifecycle.destroy();
    lifecycle.destroy();
    expect(installed.size).toBe(0);
    expect([...removals.values()]).toEqual([1, 1, 1, 1]);
  });
}

test("an empty failed replacement clears old faces once and invalidates stale completion", () => {
  const initial = { family: "Initial Face" };
  const stale = { family: "Stale Face" };
  const remove = mock((_faces: readonly { family: string }[]) => {});
  const lifecycle = new FontFaceLifecycle(remove);
  const changed = mock(() => {});
  lifecycle.subscribe(changed);
  lifecycle.beginReplacement()([initial]);
  const commitStale = lifecycle.beginReplacement();
  const commitFailure = lifecycle.beginReplacement();

  expect(changed).toHaveBeenCalledTimes(1);
  commitFailure([]);
  expect(remove).toHaveBeenCalledWith([initial]);
  expect(changed).toHaveBeenCalledTimes(2);
  commitStale([stale]);
  expect(remove).toHaveBeenCalledWith([stale]);
  expect(changed).toHaveBeenCalledTimes(2);
  lifecycle.clear();
  expect(changed).toHaveBeenCalledTimes(2);
});

test("destroy removes late completions without notifying or releasing an active face twice", () => {
  const face = { family: "Late Face" };
  const remove = mock((_faces: readonly { family: string }[]) => {});
  const lifecycle = new FontFaceLifecycle(remove);
  const changed = mock(() => {});
  lifecycle.subscribe(changed);
  const commit = lifecycle.beginReplacement();
  lifecycle.destroy();
  commit([face]);
  expect(remove).toHaveBeenCalledWith([face]);
  expect(changed).not.toHaveBeenCalled();
});
