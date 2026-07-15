/**
 * Framework-neutral lifecycle owner for persistent header/footer EditorViews.
 *
 * Each distinct relationship id owns one off-screen ProseMirror view. Painted
 * header/footer instances that share that relationship therefore share live
 * editor state, and adapters only decide when to synchronize or destroy the
 * manager.
 */

import { EditorState } from "prosemirror-state";
import type { EditorState as EditorStateT } from "prosemirror-state";
import { EditorView } from "prosemirror-view";

import { clearHeaderFooterVerbatimXml } from "../docx/headerFooterVerbatim";
import { proseDocToBlocks } from "../prosemirror/conversion/fromProseDoc";
import { headerFooterToProseDoc } from "../prosemirror/conversion/toProseDoc";
import { ExtensionManager } from "../prosemirror/extensions/ExtensionManager";
import { ensureBaseDirectionInState } from "../prosemirror/extensions/features/AutoBidiDetectionExtension";
import { ensureParaIdsInState } from "../prosemirror/extensions/features/ParaIdAllocatorExtension";
import { createStarterKit } from "../prosemirror/extensions/StarterKit";
import { createDocumentStylesPlugin } from "../prosemirror/plugins/documentStyles";
import { schema } from "../prosemirror/schema";
import type {
  BlockContent,
  Document,
  HeaderFooter,
  StyleDefinitions,
  Theme,
} from "../types/document";

export type HeaderFooterPartKind = "header" | "footer";

export type HeaderFooterPartKey = {
  kind: HeaderFooterPartKind;
  rId: string;
};

export type HeaderFooterEditorTransaction = {
  docChanged: boolean;
  kind: HeaderFooterPartKind;
  rId: string;
  selectionChanged: boolean;
  view: EditorView;
};

export type HeaderFooterEditorManagerDeps = {
  getDocument: () => Document | null;
  getHost: () => HTMLElement | null;
  getStyles: () => StyleDefinitions | null | undefined;
  getTheme: () => Theme | null | undefined;
  onTransaction?: ((transaction: HeaderFooterEditorTransaction) => void) | undefined;
};

export type HeaderFooterEditorManager = {
  destroy: () => void;
  getView: (rId: string) => EditorView | null;
  listSlots: () => HeaderFooterPartKey[];
  snapshotDocument: (document: Document) => Document;
  sync: () => void;
};

type MountedView = {
  appliedContent: BlockContent[];
  appliedHeaderFooter: HeaderFooter;
  appliedStyles: StyleDefinitions | null | undefined;
  appliedTheme: Theme | null | undefined;
  dirty: boolean;
  kind: HeaderFooterPartKind;
  manager: ExtensionManager;
  mountNode: HTMLElement;
  rId: string;
  view: EditorView;
};

const buildInitialState = (
  headerFooter: HeaderFooter,
  styles: StyleDefinitions | null | undefined,
  theme: Theme | null | undefined,
  manager: ExtensionManager,
): EditorStateT => {
  const proseDocOptions: { styles?: StyleDefinitions; theme?: Theme | null } = {};
  if (styles) {
    proseDocOptions.styles = styles;
  }
  if (theme !== undefined) {
    proseDocOptions.theme = theme;
  }

  const document = headerFooterToProseDoc(headerFooter.content, proseDocOptions);
  return ensureBaseDirectionInState(
    ensureParaIdsInState(
      EditorState.create({
        doc: document,
        schema,
        plugins: [...manager.getPlugins(), createDocumentStylesPlugin(styles)],
      }),
    ),
  );
};

export const enumerateHeaderFooterParts = ({
  headers,
  footers,
}: {
  headers: Map<string, HeaderFooter> | undefined;
  footers: Map<string, HeaderFooter> | undefined;
}): HeaderFooterPartKey[] => {
  const parts: HeaderFooterPartKey[] = [];
  if (headers) {
    for (const rId of headers.keys()) {
      parts.push({ kind: "header", rId });
    }
  }
  if (footers) {
    for (const rId of footers.keys()) {
      if (!headers?.has(rId)) {
        parts.push({ kind: "footer", rId });
      }
    }
  }
  return parts;
};

export const enumerateDocumentHeaderFooterParts = (
  document: Document | null,
): HeaderFooterPartKey[] => {
  if (!document?.package) {
    return [];
  }
  return enumerateHeaderFooterParts({
    headers: document.package.headers,
    footers: document.package.footers,
  });
};

const resolveHeaderFooter = (
  document: Document | null,
  { kind, rId }: HeaderFooterPartKey,
): HeaderFooter | null => {
  const parts = kind === "header" ? document?.package.headers : document?.package.footers;
  return parts?.get(rId) ?? null;
};

export const createHeaderFooterEditorManager = (
  deps: HeaderFooterEditorManagerDeps,
): HeaderFooterEditorManager => {
  const mounted = new Map<string, MountedView>();

  const destroyMounted = (part: MountedView): void => {
    part.view.destroy();
    part.manager.destroy();
    part.mountNode.remove();
  };

  const destroy = (): void => {
    for (const part of mounted.values()) {
      destroyMounted(part);
    }
    mounted.clear();
  };

  const sync = (): void => {
    const host = deps.getHost();
    if (!host) {
      return;
    }

    const document = deps.getDocument();
    const styles = deps.getStyles();
    const theme = deps.getTheme();
    const wanted = new Map(
      enumerateDocumentHeaderFooterParts(document).map((part) => [part.rId, part] as const),
    );

    for (const [rId, part] of mounted) {
      if (wanted.has(rId)) {
        continue;
      }
      destroyMounted(part);
      mounted.delete(rId);
    }

    for (const part of wanted.values()) {
      const headerFooter = resolveHeaderFooter(document, part);
      if (!headerFooter) {
        continue;
      }

      const existing = mounted.get(part.rId);
      if (existing) {
        if (existing.mountNode.parentElement !== host) {
          host.append(existing.mountNode);
        }
        const contentIsCurrent =
          existing.appliedHeaderFooter === headerFooter &&
          existing.appliedContent === headerFooter.content;
        const contextIsCurrent =
          existing.appliedStyles === styles && existing.appliedTheme === theme;
        if (contentIsCurrent && contextIsCurrent) {
          continue;
        }

        existing.view.updateState(buildInitialState(headerFooter, styles, theme, existing.manager));
        existing.appliedHeaderFooter = headerFooter;
        existing.appliedContent = headerFooter.content;
        existing.appliedStyles = styles;
        existing.appliedTheme = theme;
        existing.dirty = false;
        continue;
      }

      const manager = new ExtensionManager(createStarterKit());
      manager.buildSchema();
      manager.initializeRuntime();

      const mountNode = host.ownerDocument.createElement("div");
      mountNode.dataset["hfRId"] = part.rId;
      mountNode.dataset["hfKind"] = part.kind;
      host.append(mountNode);

      const view = new EditorView(mountNode, {
        state: buildInitialState(headerFooter, styles, theme, manager),
        dispatchTransaction(transaction) {
          const nextState = view.state.apply(transaction);
          view.updateState(nextState);
          const mountedPart = mounted.get(part.rId);
          if (mountedPart && transaction.docChanged) {
            mountedPart.dirty = true;
          }
          deps.onTransaction?.({
            docChanged: transaction.docChanged,
            kind: part.kind,
            rId: part.rId,
            selectionChanged: transaction.selectionSet,
            view,
          });
        },
      });

      mounted.set(part.rId, {
        appliedContent: headerFooter.content,
        appliedHeaderFooter: headerFooter,
        appliedStyles: styles,
        appliedTheme: theme,
        dirty: false,
        kind: part.kind,
        manager,
        mountNode,
        rId: part.rId,
        view,
      });
    }
  };

  return {
    destroy,
    getView: (rId) => mounted.get(rId)?.view ?? null,
    listSlots: () =>
      [...mounted.values()].map(({ kind, rId }) => ({
        kind,
        rId,
      })),
    snapshotDocument: (document) => {
      let headers = document.package.headers;
      let footers = document.package.footers;
      let headersChanged = false;
      let footersChanged = false;

      for (const { dirty, kind, rId, view } of mounted.values()) {
        if (!dirty) {
          continue;
        }
        const source = kind === "header" ? headers : footers;
        const existing = source?.get(rId);
        if (!source || !existing) {
          continue;
        }
        const updated: HeaderFooter = {
          ...existing,
          content: proseDocToBlocks(view.state.doc),
        };
        clearHeaderFooterVerbatimXml(updated);
        if (kind === "header") {
          if (!headersChanged) {
            headers = new Map(headers);
            headersChanged = true;
          }
          headers?.set(rId, updated);
        } else {
          if (!footersChanged) {
            footers = new Map(footers);
            footersChanged = true;
          }
          footers?.set(rId, updated);
        }
      }

      if (!headersChanged && !footersChanged) {
        return document;
      }
      const packageWithSnapshots = { ...document.package };
      if (headersChanged && headers) {
        packageWithSnapshots.headers = headers;
      }
      if (footersChanged && footers) {
        packageWithSnapshots.footers = footers;
      }
      return {
        ...document,
        package: packageWithSnapshots,
      };
    },
    sync,
  };
};
