import type { Document } from "@stll/folio-core";
import type { FolioEditor } from "@stll/folio-core/controller/folioEditor";

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
    return {
      active: canonical !== null && canonical !== undefined,
      document: ref?.getDocument() ?? null,
      projectionJSON: state?.doc.toJSON() ?? null,
      text: state?.doc.textContent ?? null,
      selection: state ? { from: state.selection.from, to: state.selection.to } : null,
      canUndo: editor?.canUndo() ?? false,
      canRedo: editor?.canRedo() ?? false,
    };
  },
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
