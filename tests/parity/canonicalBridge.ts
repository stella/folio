import { TextSelection, type Transaction } from "prosemirror-state";
import { applyCellSelection } from "@stll/folio-core/prosemirror/cellDragSelection";
import { singletonManager } from "@stll/folio-core/prosemirror/schema";
import type { BrowserDragTarget } from "../visual/browserDragTarget";
import { canonicalTextSelection } from "./canonicalTextSelection";
import { waitForCanonicalLoadOwner } from "./canonicalLoadOwner";
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

export type CanonicalHyperlinkAction =
  | { type: "setHyperlink"; from: number; to: number; href: string; tooltip?: string }
  | { type: "removeHyperlink"; from: number; to: number }
  | {
      type: "insertHyperlink";
      from: number;
      to: number;
      text: string;
      href: string;
      tooltip?: string;
    };

/** Private interaction-test bridge shared by both playgrounds. */
export const buildCanonicalBridge = (getRef: () => CanonicalPlaygroundRef | null) => ({
  observeInput: () => {
    const view = getRef()?.getEditor()?.getView();
    if (!view) return null;
    const record = (transaction: Transaction) => ({
      before: view.state.doc.toJSON(),
      proposed: transaction.doc.toJSON(),
      selection: view.state.selection.toJSON(),
      proposedSelection: transaction.selection.toJSON(),
      storedMarks: view.state.storedMarks?.map((mark) => mark.toJSON()) ?? null,
      steps: transaction.steps.map((step) => step.toJSON()),
      composition: transaction.getMeta("composition") ?? null,
      composing: view.composing,
    });
    const records: ReturnType<typeof record>[] = [];
    const events: (
      | { type: "beforeinput"; inputType: string; data: string | null; isComposing: boolean }
      | { type: "compositionend"; data: string }
    )[] = [];
    const controller = new AbortController();
    view.dom.addEventListener(
      "beforeinput",
      (event) => {
        events.push({
          type: "beforeinput",
          inputType: event.inputType,
          data: event.data,
          isComposing: event.isComposing,
        });
      },
      { capture: true, signal: controller.signal },
    );
    view.dom.addEventListener(
      "compositionend",
      (event) => {
        if (event instanceof CompositionEvent)
          events.push({ type: "compositionend", data: event.data });
      },
      { capture: true, signal: controller.signal },
    );
    const dispatch = view.dispatch;
    view.dispatch = (transaction) => {
      records.push(record(transaction));
      dispatch(transaction);
    };
    return () => {
      view.dispatch = dispatch;
      controller.abort();
      return { transactions: records, events };
    };
  },
  nativeComposing: () => getRef()?.getEditor()?.getView()?.composing ?? null,
  ensureView: () => {
    const ref = getRef();
    if (!ref) return false;
    ref.ensureEditorView();
    return true;
  },
  load: async (bytes: number[]) => {
    const ref = getRef();
    if (!ref) return false;
    const previousOwner = ref.getEditor()?.getCanonicalDocument();
    await ref.loadDocumentBuffer(new Uint8Array(bytes));
    ref.ensureEditorView();
    await waitForCanonicalLoadOwner({
      previousOwner,
      getOwner: () => getRef()?.getEditor()?.getCanonicalDocument(),
      waitFrame: () => new Promise<void>((resolve) => requestAnimationFrame(() => resolve())),
    });
    return true;
  },
  save: async () => {
    const buffer = await getRef()?.save();
    return buffer ? [...new Uint8Array(buffer)] : null;
  },
  canSnapshot: () => getRef()?.getEditor()?.getCanonicalStoryProjection("main") != null,
  committedVersion: () => getRef()?.getEditor()?.captureCanonicalSave()?.version ?? null,
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
        return token === undefined
          ? getProseParagraphPropertySourceToken(node) == null &&
              getParagraphPropertySource(block) === undefined
          : paragraphPropertySourceTokenMatchesContract(token, sourceContract) &&
              getProseParagraphPropertySourceToken(node) === token &&
              (getParagraphPropertySource(block) === undefined ||
                paragraphPropertySourceBelongsToDocument(block, canonical));
      }),
    );
    return {
      active: canonical !== null && canonical !== undefined,
      composing: editor?.getView()?.composing ?? null,
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
      selectionJSON: state?.selection.toJSON() ?? null,
      canUndo: editor?.canUndo() ?? false,
      canRedo: editor?.canRedo() ?? false,
    };
  },
  selectStructuralTarget: (target: BrowserDragTarget) => {
    const view = getRef()?.getEditor()?.getView();
    if (!view) return null;
    const targets: { pos: number; size: number; type: "paragraph" | "inline" }[] = [];
    const name = {
      table: "tableCell",
      list: "paragraph",
      note: "footnoteRef",
      field: "field",
      inlineObject: "image",
    }[target];
    view.state.doc.descendants((node, pos) => {
      if (targets.length >= 2) return false;
      const matches = node.type.name === name && (target !== "list" || node.attrs["numPr"] != null);
      if (matches || (target === "note" && node.marks.some((mark) => mark.type.name === name)))
        targets.push({
          pos,
          size: node.nodeSize,
          type: node.type.name === "paragraph" ? "paragraph" : "inline",
        });
      return target !== "table" || !matches;
    });
    const first = targets.at(0);
    const last = targets.at(1) ?? first;
    if (!first || !last) return null;
    if (target === "table") {
      if (!applyCellSelection(view, first.pos, last.pos)) return null;
    } else {
      const from = Math.max(
        view.state.doc.resolve(first.pos).start(),
        first.type === "paragraph" ? first.pos + 1 : first.pos - 1,
      );
      const to = Math.min(
        view.state.doc.content.size - 1,
        last.type === "paragraph"
          ? last.pos + last.size - 1
          : view.state.doc.resolve(last.pos).end(),
        last.type === "paragraph"
          ? last.pos + Math.min(2, last.size - 1)
          : last.pos + last.size + 2,
      );
      const selection = TextSelection.create(view.state.doc, from, to);
      view.dispatch(view.state.tr.setSelection(selection));
    }
    view.focus();
    return view.state.selection.toJSON();
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
  executeHyperlink: (action: CanonicalHyperlinkAction) => {
    const editor = getRef()?.getEditor();
    if (!editor) return false;
    editor.setSelection(action.from, action.to);
    switch (action.type) {
      case "setHyperlink":
        return editor.executeCommand(
          singletonManager.requireCommand("setHyperlink")(action.href, action.tooltip),
        );
      case "removeHyperlink":
        return editor.executeCommand(singletonManager.requireCommand("removeHyperlink")());
      case "insertHyperlink":
        return editor.executeCommand(
          singletonManager.requireCommand("insertHyperlink")(
            action.text,
            action.href,
            action.tooltip,
          ),
        );
      default: {
        const unreachable: never = action;
        return unreachable;
      }
    }
  },
});

declare global {
  var __folioCanonicalReady: boolean | undefined;
  var __folioCanonical: ReturnType<typeof buildCanonicalBridge> | undefined;
}
