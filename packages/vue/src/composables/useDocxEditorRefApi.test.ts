import { describe, expect, mock, test } from "bun:test";
import { history, undo } from "prosemirror-history";
import { EditorState, type Transaction } from "prosemirror-state";

import {
  commentRecordsOf,
  createCommentRecordsPlugin,
} from "@stll/folio-core/prosemirror/commentRecords";
import { schema } from "@stll/folio-core/prosemirror/schema";

import {
  applyWithStagedOperationComments,
  saveDocumentForHost,
} from "./useDocxEditorRefApi";

/** A view whose state holds no comment records: the host keeps the list. */
const viewWithoutRecords = () => ({
  state: EditorState.create({ schema }),
  dispatch: () => undefined,
});

describe("applyWithStagedOperationComments", () => {
  test("publishes only applied comments after the document commits", () => {
    const events: string[] = [];
    let nextId = 0;
    const result = applyWithStagedOperationComments({
      view: viewWithoutRecords(),
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
    expect(events).toEqual([
      "mint:refused",
      "mint:accepted",
      "commit",
      "publish:2",
    ]);
  });

  test("records the comments the operation anchors in its own transaction", () => {
    const doc = schema.node("doc", null, [
      schema.node("paragraph", null, schema.text("Clause")),
    ]);
    const holder = {
      state: EditorState.create({
        doc,
        plugins: [history(), createCommentRecordsPlugin([])],
      }),
    };
    const view = {
      get state() {
        return holder.state;
      },
      dispatch: (transaction: Transaction) => {
        holder.state = holder.state.apply(transaction);
      },
    };
    const publishComments = mock(() => {});
    let nextId = 40;
    applyWithStagedOperationComments({
      view,
      createComment: (text) => ({
        id: ++nextId,
        author: "Reviewer",
        content: [
          {
            type: "paragraph",
            content: [{ type: "run", content: [{ type: "text", text }] }],
          },
        ],
      }),
      publishComments,
      apply: (createCommentId, recordingView) => {
        createCommentId("refused");
        const accepted = createCommentId("accepted");
        recordingView.dispatch(
          recordingView.state.tr.addMark(
            1,
            7,
            schema.marks.comment.create({ commentId: accepted }),
          ),
        );
        return { applied: [{ commentId: accepted }] };
      },
    });
    expect(publishComments).not.toHaveBeenCalled();
    expect(commentRecordsOf(holder.state)?.map(({ id }) => id)).toEqual([42]);
    // One undo takes the mark and the record back together.
    undo(holder.state, view.dispatch);
    expect(commentRecordsOf(holder.state)).toEqual([]);
    expect(holder.state.doc.eq(doc)).toBe(true);
  });

  test("does not notify the host when every operation is refused", () => {
    const publishComments = mock(() => {});
    applyWithStagedOperationComments({
      view: viewWithoutRecords(),
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
