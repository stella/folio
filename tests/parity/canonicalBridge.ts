import { canonicalTextSelection } from "./canonicalTextSelection";
import type { Document } from "@stll/folio-core";
import type { FolioEditor } from "@stll/folio-core/controller/folioEditor";
import { toProseDoc } from "@stll/folio-core/prosemirror/conversion/toProseDoc";
import {
  getDocumentParagraphPropertySourceContract,
  getParagraphPropertySource,
  getParagraphPropertySourceToken,
  getProseDocumentParagraphPropertySourceContract,
  getProseParagraphPropertySourceToken,
  paragraphPropertySourceBelongsToDocument,
  paragraphPropertySourceTokenMatchesContract,
} from "../../packages/core/src/docx/paragraphPropertySource";

type CanonicalPlaygroundRef = {
  getDocument: () => Document | null;
  getEditor: () => FolioEditor | null;
  ensureEditorView: () => void;
  loadDocumentBuffer: (buffer: Uint8Array) => Promise<void>;
  save: () => Promise<ArrayBuffer | null>;
};

/** Private interaction-test bridge shared by both playgrounds. */
export const buildCanonicalBridge = (getRef: () => CanonicalPlaygroundRef | null) => ({
  load: async (bytes: number[]) => {
    const ref = getRef();
    if (!ref) return false;
    await ref.loadDocumentBuffer(new Uint8Array(bytes));
    ref.ensureEditorView();
    return true;
  },
  save: async () => {
    const buffer = await getRef()?.save();
    return buffer ? [...new Uint8Array(buffer)] : null;
  },
  snapshot: () => {
    const ref = getRef();
    const editor = ref?.getEditor();
    const state = editor?.getState();
    const canonical = editor?.getCanonicalDocument();
    // Projection and private source ownership must be checked before JSON transport drops the bindings.
    const canonicalProjection = canonical
      ? toProseDoc(
          canonical,
          canonical.package.styles ? { styles: canonical.package.styles } : undefined,
        )
      : null;
    const sourceContract = canonical
      ? getDocumentParagraphPropertySourceContract(canonical)
      : undefined;
    const content = canonical?.package.document.content ?? [];
    const capturedParagraphCount = content.filter(
      (block) => block.type === "paragraph" && getParagraphPropertySource(block) !== undefined,
    ).length;
    const provenanceValid = Boolean(
      canonical &&
      state &&
      sourceContract &&
      sourceContract === getProseDocumentParagraphPropertySourceContract(state.doc) &&
      content.length > 0 &&
      content.length === state.doc.childCount &&
      content.every((block, index) => {
        if (block.type !== "paragraph") return false;
        const token = getParagraphPropertySourceToken(block);
        const node = state.doc.child(index);
        return (
          paragraphPropertySourceTokenMatchesContract(token, sourceContract) &&
          getProseParagraphPropertySourceToken(node) === token &&
          (getParagraphPropertySource(block) === undefined ||
            paragraphPropertySourceBelongsToDocument(block, canonical))
        );
      }),
    );
    return {
      active: canonical !== null && canonical !== undefined,
      document: ref?.getDocument() ?? null,
      projectionJSON: state?.doc.toJSON() ?? null,
      canonicalProjectionJSON: canonicalProjection?.toJSON() ?? null,
      projectionMatchesCanonical: Boolean(
        state && canonicalProjection && state.doc.eq(canonicalProjection),
      ),
      provenance: { valid: provenanceValid, capturedParagraphCount },
      text: state?.doc.textContent ?? null,
      textSelection: state ? canonicalTextSelection(state) : null,
      selection: state ? { from: state.selection.from, to: state.selection.to } : null,
      canUndo: editor?.canUndo() ?? false,
      canRedo: editor?.canRedo() ?? false,
    };
  },
  setMode: (mode: "editing" | "suggesting") =>
    getRef()
      ?.getEditor()
      ?.setCanonicalMode(
        mode === "suggesting"
          ? { type: "suggesting", author: "Canonical test author" }
          : { type: "editing" },
      ) ?? false,
  resolveRevisions: (revisionIds: readonly number[], resolution: "accept" | "reject") =>
    getRef()?.getEditor()?.resolveCanonicalRevisions(revisionIds, resolution) ?? false,
  select: (from: number, to = from) => {
    const editor = getRef()?.getEditor();
    if (!editor?.getView()) return false;
    editor.setSelection(from, to);
    editor.focus();
    return true;
  },
});

declare global {
  var __folioCanonical: ReturnType<typeof buildCanonicalBridge> | undefined;
}
