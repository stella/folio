import { describe, expect, mock, test } from "bun:test";

import { applyWithStagedOperationComments, saveDocumentForHost } from "./useDocxEditorRefApi";

describe("applyWithStagedOperationComments", () => {
  test("publishes only applied comments after the document commits", () => {
    const events: string[] = [];
    let nextId = 0;
    const result = applyWithStagedOperationComments({
      createComment: (text) => {
        const id = ++nextId;
        events.push(`mint:${text}`);
        return { id, author: "Reviewer", content: [] };
      },
      publishComments: (comments) => {
        events.push(`publish:${comments.map(({ id }) => id).join(",")}`);
      },
      apply: (createCommentId) => {
        createCommentId("refused");
        const accepted = createCommentId("accepted");
        events.push("commit");
        return { applied: [{ commentId: accepted }] };
      },
    });

    expect(result.applied).toEqual([{ commentId: 2 }]);
    expect(events).toEqual(["mint:refused", "mint:accepted", "commit", "publish:2"]);
  });

  test("does not notify the host when every operation is refused", () => {
    const publishComments = mock(() => {});
    applyWithStagedOperationComments({
      createComment: () => ({ id: 1, author: "Reviewer", content: [] }),
      publishComments,
      apply: (createCommentId) => {
        createCommentId("refused");
        return { applied: [] };
      },
    });
    expect(publishComments).not.toHaveBeenCalled();
  });
});

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
