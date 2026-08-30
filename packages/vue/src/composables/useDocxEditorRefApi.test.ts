import { describe, expect, mock, test } from "bun:test";

import { saveDocumentForHost } from "./useDocxEditorRefApi";

describe("saveDocumentForHost", () => {
  test("runs host-facing effects only after serialization succeeds", async () => {
    const bytes = new Uint8Array([1, 2, 3]);
    const clearCommentsDirty = mock(() => {});
    const onSave = mock((_buffer: ArrayBuffer) => {});
    const saveDocument = mock((_options?: { selective?: boolean }) =>
      Promise.resolve(new Blob([bytes])),
    );

    const result = await saveDocumentForHost(
      { clearCommentsDirty, onSave, saveDocument },
      { selective: false },
    );

    expect(saveDocument).toHaveBeenCalledWith({ selective: false });
    expect(new Uint8Array(result ?? new ArrayBuffer(0))).toEqual(bytes);
    expect(clearCommentsDirty).toHaveBeenCalledTimes(1);
    expect(onSave).toHaveBeenCalledWith(result);
  });

  test("does not run host-facing effects when serialization has no document", async () => {
    const clearCommentsDirty = mock(() => {});
    const onSave = mock((_buffer: ArrayBuffer) => {});

    const result = await saveDocumentForHost({
      clearCommentsDirty,
      onSave,
      saveDocument: () => Promise.resolve(null),
    });

    expect(result).toBeNull();
    expect(clearCommentsDirty).not.toHaveBeenCalled();
    expect(onSave).not.toHaveBeenCalled();
  });
});
