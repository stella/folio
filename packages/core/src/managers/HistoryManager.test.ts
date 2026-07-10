import { describe, expect, test } from "bun:test";

import { HistoryManager, classifyHistoryShortcut } from "./HistoryManager";

describe("HistoryManager", () => {
  test("push records undo entries and exposes the current state", () => {
    const manager = new HistoryManager<number>(0, { groupingInterval: 0 });
    manager.push(1);
    manager.push(2);

    expect(manager.state).toBe(2);
    expect(manager.getSnapshot()).toMatchObject({
      state: 2,
      canUndo: true,
      canRedo: false,
      undoCount: 2,
      redoCount: 0,
    });
  });

  test("skips no-op pushes via the equality check", () => {
    const manager = new HistoryManager<number>(0, { groupingInterval: 0 });
    manager.push(0);
    expect(manager.getSnapshot().undoCount).toBe(0);
  });

  test("undo and redo move between recorded states", () => {
    const manager = new HistoryManager<number>(0, { groupingInterval: 0 });
    manager.push(1);
    manager.push(2);

    expect(manager.undo()).toBe(1);
    expect(manager.state).toBe(1);
    expect(manager.getSnapshot()).toMatchObject({ canUndo: true, canRedo: true, redoCount: 1 });

    expect(manager.redo()).toBe(2);
    expect(manager.state).toBe(2);
    expect(manager.getSnapshot().canRedo).toBe(false);
  });

  test("a new push after undo invalidates the redo stack", () => {
    const manager = new HistoryManager<number>(0, { groupingInterval: 0 });
    manager.push(1);
    manager.push(2);
    manager.undo();

    expect(manager.getSnapshot().redoCount).toBe(1);
    // A real edit only records once the guard is lowered (the adapter does this
    // after its render settles); that push clears the redo stack.
    manager.endUndoRedo();
    manager.push(3);
    expect(manager.getSnapshot().redoCount).toBe(0);
    expect(manager.redo()).toBeUndefined();
  });

  test("groups rapid pushes into one undo entry; undo restores to the group start", () => {
    // A wide window groups every push that lands within it.
    const manager = new HistoryManager<number>(0, { groupingInterval: 10_000 });
    manager.push(1);
    manager.push(2);
    manager.push(3);

    // The rapid run collapses to a single undo entry that keeps the state from
    // before the group, so one undo steps back past the whole group (1,2,3 -> 0),
    // not just the last push.
    expect(manager.getSnapshot().undoCount).toBe(1);
    expect(manager.undo()).toBe(0);
  });

  test("caps the undo stack at maxEntries", () => {
    const manager = new HistoryManager<number>(0, { groupingInterval: 0, maxEntries: 3 });
    for (let i = 1; i <= 5; i++) {
      manager.push(i);
    }
    expect(manager.getSnapshot().undoCount).toBe(3);
  });

  test("the re-entrancy guard suppresses entry creation until lowered", () => {
    const manager = new HistoryManager<number>(0, { groupingInterval: 0 });
    manager.push(1);
    manager.undo();
    expect(manager.isUndoRedoActive).toBe(true);

    // A push that arrives while the guard is up adopts the state without
    // recording an entry or touching the redo stack.
    manager.push(99);
    expect(manager.state).toBe(99);
    expect(manager.getSnapshot().undoCount).toBe(0);
    expect(manager.getSnapshot().redoCount).toBe(1);

    manager.endUndoRedo();
    expect(manager.isUndoRedoActive).toBe(false);
    manager.push(100);
    expect(manager.getSnapshot().undoCount).toBe(1);
  });

  test("reset returns to the initial state and clears history", () => {
    const manager = new HistoryManager<number>(7, { groupingInterval: 0 });
    manager.push(1);
    manager.push(2);
    manager.reset();

    expect(manager.state).toBe(7);
    expect(manager.getSnapshot()).toMatchObject({ canUndo: false, canRedo: false });

    manager.reset(42);
    expect(manager.state).toBe(42);
  });

  test("transformAll rewrites the current state and both stacks", () => {
    const manager = new HistoryManager<number>(0, { groupingInterval: 0 });
    manager.push(1);
    manager.push(2);
    manager.undo();

    manager.transformAll((value) => value * 10);
    expect(manager.state).toBe(10);
    expect(manager.getUndoStack().map((entry) => entry.state)).toEqual([0]);
    expect(manager.getRedoStack().map((entry) => entry.state)).toEqual([20]);
  });

  test("satisfies the DocumentLoaderHistory shape", () => {
    const manager = new HistoryManager<{ id: number } | null>(null);
    const loaderView: {
      readonly state: { id: number } | null;
      reset: (doc: { id: number }) => void;
    } = manager;
    loaderView.reset({ id: 1 });
    expect(loaderView.state).toEqual({ id: 1 });
  });
});

describe("classifyHistoryShortcut", () => {
  const event = (key: string, mods: { ctrl?: boolean; meta?: boolean; shift?: boolean } = {}) => ({
    key,
    ctrlKey: mods.ctrl ?? false,
    metaKey: mods.meta ?? false,
    shiftKey: mods.shift ?? false,
  });

  test("Ctrl/Cmd+Z is undo", () => {
    expect(classifyHistoryShortcut(event("z", { ctrl: true }))).toBe("undo");
    expect(classifyHistoryShortcut(event("z", { meta: true }))).toBe("undo");
  });

  test("Ctrl/Cmd+Y is redo", () => {
    expect(classifyHistoryShortcut(event("y", { ctrl: true }))).toBe("redo");
    expect(classifyHistoryShortcut(event("Y", { meta: true }))).toBe("redo");
  });

  test("Ctrl/Cmd+Shift+Z is redo even though Shift reports the uppercase key", () => {
    // The browser delivers `event.key === "Z"` (uppercase) while Shift is held;
    // a case-sensitive `=== "z"` check would miss the standard redo chord.
    expect(classifyHistoryShortcut(event("Z", { meta: true, shift: true }))).toBe("redo");
    expect(classifyHistoryShortcut(event("Z", { ctrl: true, shift: true }))).toBe("redo");
    // Defensive: a lowercase "z" with Shift (some synthetic event sources) too.
    expect(classifyHistoryShortcut(event("z", { meta: true, shift: true }))).toBe("redo");
  });

  test("ignores chords without a Ctrl/Cmd modifier or unrelated keys", () => {
    expect(classifyHistoryShortcut(event("z"))).toBeNull();
    expect(classifyHistoryShortcut(event("a", { meta: true }))).toBeNull();
    expect(classifyHistoryShortcut(event("Z", { shift: true }))).toBeNull();
  });
});
