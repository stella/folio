import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { panic } from "better-result";
import { afterAll, expect, test } from "bun:test";

import type { Comment } from "@stll/folio-core/types/content";

// Vue reads the DOM globals at import time, so it loads after registration.
const { createApp, defineComponent, h, ref } = await import("vue");
const { useCommentManagement } = await import("./useCommentManagement");

afterAll(() => {
  GlobalRegistrator.unregister();
});

const THREAD: Comment = {
  id: 4,
  author: "Owner",
  content: [
    {
      type: "paragraph",
      content: [{ type: "run", content: [{ type: "text", text: "Check." }] }],
    },
  ],
};

type Harness = {
  management: ReturnType<typeof useCommentManagement>;
  changes: Comment[][];
  refusals: number;
  unmount: () => void;
};

const mountManagement = (readOnly: boolean): Harness => {
  const holder: { management: ReturnType<typeof useCommentManagement> | null } =
    { management: null };
  const state = { changes: [] as Comment[][], refusals: 0 };
  const container = document.createElement("div");
  const app = createApp(
    defineComponent({
      setup() {
        holder.management = useCommentManagement({
          editorView: ref(null),
          getDocument: () => null,
          author: () => "Viewer",
          commentsProp: () => [THREAD],
          onCommentsChange: (next) => {
            state.changes.push(next);
          },
          reLayout: () => undefined,
          readOnly: () => readOnly,
          onReadOnlyEditAttempt: () => {
            state.refusals += 1;
          },
        });
        return () => h("div");
      },
    }),
  );
  app.mount(container);
  return {
    management: holder.management ?? panic("Expected comment management"),
    get changes() {
      return state.changes;
    },
    get refusals() {
      return state.refusals;
    },
    unmount: () => app.unmount(),
  };
};

test("a read-only editor publishes no comment thread change", () => {
  const harness = mountManagement(true);
  try {
    const { management } = harness;
    management.handleResolve(THREAD.id);
    management.handleUnresolve(THREAD.id);
    management.handleReply(THREAD.id, "Reply");
    management.handleTrackedChangeReply(11, "Reply");
    management.handleDelete(THREAD.id);
    expect(harness.changes).toEqual([]);
    expect(management.commentsDirty.value).toBe(false);
    expect(harness.refusals).toBe(5);
  } finally {
    harness.unmount();
  }
});

test("without a canonical session a comment thread change reports the document read-only", () => {
  const harness = mountManagement(false);
  try {
    const { management } = harness;
    management.handleResolve(THREAD.id);
    management.handleDelete(THREAD.id);
    expect(harness.changes).toEqual([]);
    expect(management.commentsDirty.value).toBe(false);
    expect(harness.refusals).toBe(2);
  } finally {
    harness.unmount();
  }
});
