import type { DocxEditorRef } from "../../packages/react/src/components/DocxEditor.props";

declare global {
  var __folioPlayground: { getEditorRef: () => DocxEditorRef | null } | undefined;
  var __folioInterleavingSuggest: ((text: string) => number) | undefined;
}

type InterleavingRef = Pick<
  DocxEditorRef,
  "createAIEditSnapshot" | "applyDocumentOperations" | "getTrackedChanges"
>;

/** Apply model-authored edits through the live browser document-operation API. */
export const createInterleavingSuggest = (ref: InterleavingRef) => {
  let operationSequence = 0;
  return (text: string) => {
    const snapshot = ref.createAIEditSnapshot();
    if (!snapshot) throw new Error("interleaving editor snapshot unavailable");
    const block = snapshot.blocks.find(
      (candidate) => !candidate.table && candidate.kind !== "diagnostic",
    );
    if (!block) throw new Error("interleaving fixture has no body block");
    const blockTextHash = snapshot.anchors[block.id]?.textHash;
    if (!blockTextHash) throw new Error("interleaving block precondition unavailable");
    operationSequence++;
    const id = `interleaving-${operationSequence}`;
    // Preserve the previous agent bridge's tracked-change mode and text-hash
    // guard. Tool argument parsing remains covered by the agent package tests.
    const options = {
      snapshot,
      author: "Fuzz reviewer",
      mode: "tracked-changes",
      batch: {
        version: 1,
        mode: "tracked-changes",
        operations: [
          {
            id,
            type: "insertAfterBlock",
            blockId: block.id,
            text,
            precondition: { blockTextHash },
          },
        ],
      },
    } as const satisfies Parameters<DocxEditorRef["applyDocumentOperations"]>[0];
    const result = ref.applyDocumentOperations(options);
    if (
      result.status !== "committed" ||
      result.applied.length !== 1 ||
      result.skipped.length !== 0 ||
      result.applied.at(0)?.id !== id
    ) {
      throw new Error(`interleaving edit did not apply: ${JSON.stringify(result)}`);
    }
    return ref.getTrackedChanges().length;
  };
};

/** Browser-loaded module; keep server-only agent dispatch outside this graph. */
export const installInterleavingBridge = () => {
  const ref = globalThis.__folioPlayground?.getEditorRef();
  if (!ref) throw new Error("interleaving editor unavailable");
  globalThis.__folioInterleavingSuggest = createInterleavingSuggest(ref);
};
