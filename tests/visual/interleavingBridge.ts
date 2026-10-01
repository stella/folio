import { createEditorRefBridge } from "../../packages/agents/src/bridges/editor-ref";
import { executeFolioToolCall } from "../../packages/agents/src/execute";
import { FOLIO_AGENT_TOOL_NAMES } from "../../packages/agents/src/types";
import type { DocxEditorRef } from "../../packages/react/src/components/DocxEditor.props";

declare global {
  var __folioPlayground: { getEditorRef: () => DocxEditorRef | null } | undefined;
  var __folioInterleavingSuggest: ((text: string) => number) | undefined;
}

/** Browser-loaded test module: drive the same tool boundary an integrated host uses. */
export const installInterleavingBridge = () => {
  const ref = globalThis.__folioPlayground?.getEditorRef();
  if (!ref) throw new Error("interleaving editor unavailable");
  const bridge = createEditorRefBridge({
    ref,
    author: "Fuzz reviewer",
    getComments: () => [],
    setComments: () => {
      throw new Error("interleaving trace cannot write comments");
    },
  });
  globalThis.__folioInterleavingSuggest = (text) => {
    const block = bridge
      .snapshot()
      .blocks.find((candidate) => !candidate.table && candidate.kind !== "diagnostic");
    if (!block) throw new Error("interleaving fixture has no body block");
    const result = executeFolioToolCall(
      FOLIO_AGENT_TOOL_NAMES.suggestChanges,
      {
        operations: [{ type: "insertAfterBlock", blockId: block.id, text }],
      },
      bridge,
    );
    if (!result.ok) throw new Error(`suggest_changes failed: ${result.error}`);
    if (result.result.applied.length !== 1 || result.result.skipped.length !== 0) {
      throw new Error(`suggest_changes did not apply: ${JSON.stringify(result.result)}`);
    }
    return ref.getTrackedChanges().length;
  };
};
