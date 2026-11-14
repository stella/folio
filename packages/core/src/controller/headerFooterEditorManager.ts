import type { Node as PMNode } from "prosemirror-model";
import { EditorState } from "prosemirror-state";
/**
 * Framework-neutral lifecycle owner for persistent header/footer EditorViews.
 *
 * Each distinct relationship id owns one off-screen ProseMirror view. Painted
 * header/footer instances that share that relationship therefore share live
 * editor state, and adapters only decide when to synchronize or destroy the
 * manager.
 */
import type { EditorState as EditorStateT } from "prosemirror-state";
import { type DirectEditorProps, EditorView } from "prosemirror-view";

import { OP_STORIES } from "@stll/docx-core/ops";

import { CanonicalSessionError } from "../canonical/canonicalSession";
import { clearHeaderFooterVerbatimXml } from "../docx/headerFooterVerbatim";
import { cloneParagraphWithPropertySource } from "../docx/paragraphPropertySource";
import { proseDocToBlocks } from "../prosemirror/conversion/fromProseDoc";
import { storyConversionOptions } from "../prosemirror/conversion/storyConversionOptions";
import {
  headerFooterToProseDoc,
  type ToProseDocOptions,
} from "../prosemirror/conversion/toProseDoc";
import { ExtensionManager } from "../prosemirror/extensions/ExtensionManager";
import { ensureBaseDirectionInState } from "../prosemirror/extensions/features/AutoBidiDetectionExtension";
import { ensureParaIdsInState } from "../prosemirror/extensions/features/ParaIdAllocatorExtension";
import { createStarterKit } from "../prosemirror/extensions/StarterKit";
import { createNativeSelectionRefreshPlugin } from "../prosemirror/nativeSelection";
import { noteReferenceTransactionIssue } from "../prosemirror/noteReferenceOccurrences";
import { createDocumentNumberingPlugin } from "../prosemirror/plugins/documentNumbering";
import {
  createDocumentStyleContextPlugin,
  createDocumentStylesPlugin,
} from "../prosemirror/plugins/documentStyles";
import { schema } from "../prosemirror/schema";
import { storyListNumbering } from "../prosemirror/storyListNumbering";
import {
  CANONICAL_GAP,
  type CanonicalRefusalGround,
  type EditorSession,
  usesCanonicalSession,
} from "../types/canonicalCapabilities";
import type {
  BlockContent,
  Document,
  HeaderFooter,
  NumberingDefinitions,
  StyleDefinitions,
  Theme,
} from "../types/document";
import {
  type DeepReadonly,
  readModel,
  type ReadonlyDocument,
} from "../types/documentModel";
import {
  createCanonicalStoryEditor,
  refreshCanonicalStoryView,
} from "./canonicalStoryEditor";
import type { HiddenEditorApi } from "./hiddenEditorApi";

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
  createView?:
    | ((mountNode: HTMLElement, props: DirectEditorProps) => EditorView)
    | undefined;
  getDocument: () => ReadonlyDocument | null;
  getReadOnly?: (() => boolean) | undefined;
  getCanonicalApi?: (() => HiddenEditorApi | null) | undefined;
  /**
   * The canonical session's own document, read without a copy (see
   * `HiddenEditorManager.readCanonicalDocument`); null outside a session.
   */
  readCanonicalDocument: () => ReadonlyDocument | null;
  getExperimentalSession?: (() => EditorSession | undefined) | undefined;
  onSessionRefusal?:
    | ((reason: string, ground: CanonicalRefusalGround, error?: Error) => void)
    | undefined;
  getHost: () => HTMLElement | null;
  getStyles: () => DeepReadonly<StyleDefinitions> | null | undefined;
  getTheme: () => DeepReadonly<Theme> | null | undefined;
  onTransaction?:
    ((transaction: HeaderFooterEditorTransaction) => void) | undefined;
};

const createEditorView = (
  mountNode: HTMLElement,
  props: DirectEditorProps,
): EditorView => new EditorView(mountNode, props);

export type HeaderFooterEditorManager = {
  destroy: () => void;
  getView: (rId: string) => EditorView | null;
  listSlots: () => HeaderFooterPartKey[];
  /**
   * The model with the edited parts written back. It shares what no edit
   * touched with `model` and with the mounted editors, so it is read-only: a
   * writer copies it first (`cloneDocumentWithParagraphPropertySources`).
   */
  snapshotDocument: (model: ReadonlyDocument) => ReadonlyDocument;
  sync: () => void;
};

type MountedView = {
  appliedContent: BlockContent[];
  appliedHeaderFooter: HeaderFooter;
  appliedNumbering: NumberingDefinitions | null | undefined;
  appliedStyles: DeepReadonly<StyleDefinitions> | null | undefined;
  appliedTheme: DeepReadonly<Theme> | null | undefined;
  dirty: boolean;
  kind: HeaderFooterPartKind;
  manager: ExtensionManager;
  mountNode: HTMLElement;
  rId: string;
  view: EditorView;
};

const DETACHED_WATERMARK_HOST = Symbol.for("stll.detachedWatermarkHost");

const headerFooterToProseDocWithDetachedWatermarkHost = (
  headerFooter: HeaderFooter,
  options: ToProseDocOptions,
): PMNode => {
  const markedContent: BlockContent[] = headerFooter.content.map(
    (block, blockIndex) => {
      if (
        blockIndex !== headerFooter.watermarkBlockIndex ||
        block.type !== "paragraph"
      ) {
        return block;
      }
      const marked = cloneParagraphWithPropertySource(block, {});
      Reflect.set(marked, DETACHED_WATERMARK_HOST, true);
      return marked;
    },
  );
  return headerFooterToProseDoc(markedContent, options);
};

const buildInitialState = (
  headerFooter: HeaderFooter,
  styles: DeepReadonly<StyleDefinitions> | null | undefined,
  theme: DeepReadonly<Theme> | null | undefined,
  numbering: NumberingDefinitions | null | undefined,
  manager: ExtensionManager,
): EditorStateT => {
  const proseDocOptions = storyConversionOptions({ styles, theme, numbering });
  const document =
    headerFooter.watermarkBlockIndex === undefined
      ? headerFooterToProseDoc(headerFooter.content, proseDocOptions)
      : headerFooterToProseDocWithDetachedWatermarkHost(
          headerFooter,
          proseDocOptions,
        );
  return ensureBaseDirectionInState(
    ensureParaIdsInState(
      EditorState.create({
        doc: document,
        schema,
        plugins: [
          createDocumentNumberingPlugin(numbering),
          ...manager.getPlugins(),
          createDocumentStylesPlugin(styles),
        ],
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
  const parts =
    kind === "header" ? document?.package.headers : document?.package.footers;
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

  const getOwnedDocument = (): ReadonlyDocument | null =>
    usesCanonicalSession(
      deps.getExperimentalSession?.(),
      CANONICAL_GAP.authorityRouting,
    )
      ? deps.readCanonicalDocument()
      : deps.getDocument();

  const sync = (): void => {
    const host = deps.getHost();
    if (!host) {
      return;
    }

    const getReadOnly = deps.getReadOnly;
    if (getReadOnly) {
      const readOnly = getReadOnly();
      for (const part of mounted.values()) {
        const expectedEditable = !readOnly;
        if (part.view.editable === expectedEditable) {
          continue;
        }
        part.view.setProps({
          editable: () => !getReadOnly(),
        });
      }
    }

    // Pending IME input owns its view until the shared session commits.
    if (
      usesCanonicalSession(
        deps.getExperimentalSession?.(),
        CANONICAL_GAP.authorityRouting,
      ) &&
      !deps.getCanonicalApi?.()?.getCanonicalStoryProjection(OP_STORIES.MAIN)
    )
      return;

    const owned = getOwnedDocument();
    const document = owned && readModel(owned);
    const styles = deps.getStyles();
    const theme = deps.getTheme();
    const numbering = document?.package.numbering;
    const wanted = new Map(
      enumerateDocumentHeaderFooterParts(document).map(
        (part) => [part.rId, part] as const,
      ),
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
        if (
          usesCanonicalSession(
            deps.getExperimentalSession?.(),
            CANONICAL_GAP.authorityRouting,
          )
        ) {
          refreshCanonicalStoryView({
            view: existing.view,
            story: { kind: part.kind, rId: part.rId },
            api: deps.getCanonicalApi?.(),
          });
          existing.appliedHeaderFooter = headerFooter;
          existing.appliedContent = headerFooter.content;
          existing.appliedStyles = styles;
          existing.appliedTheme = theme;
          existing.appliedNumbering = numbering;
          existing.dirty = false;
          continue;
        }
        const contentIsCurrent =
          existing.appliedHeaderFooter === headerFooter &&
          existing.appliedContent === headerFooter.content;
        const contextIsCurrent =
          existing.appliedStyles === styles &&
          existing.appliedTheme === theme &&
          existing.appliedNumbering === numbering;
        if (contentIsCurrent && contextIsCurrent) {
          continue;
        }

        existing.view.updateState(
          buildInitialState(
            headerFooter,
            styles,
            theme,
            numbering,
            existing.manager,
          ),
        );
        existing.appliedHeaderFooter = headerFooter;
        existing.appliedContent = headerFooter.content;
        existing.appliedStyles = styles;
        existing.appliedTheme = theme;
        existing.appliedNumbering = numbering;
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

      const canonical = createCanonicalStoryEditor({
        story: { kind: part.kind, rId: part.rId },
        getView: () => view,
        getApi: () => deps.getCanonicalApi?.() ?? null,
        enabled: () =>
          usesCanonicalSession(
            deps.getExperimentalSession?.(),
            CANONICAL_GAP.authorityRouting,
          ),
        onRefusal: deps.onSessionRefusal,
        onSelectionChange: () =>
          deps.onTransaction?.({
            kind: part.kind,
            rId: part.rId,
            view,
            docChanged: false,
            selectionChanged: true,
          }),
      });
      const canonicalProjection = usesCanonicalSession(
        deps.getExperimentalSession?.(),
        CANONICAL_GAP.authorityRouting,
      )
        ? deps
            .getCanonicalApi?.()
            ?.getCanonicalStoryProjection({ kind: part.kind, rId: part.rId })
        : null;
      if (
        usesCanonicalSession(
          deps.getExperimentalSession?.(),
          CANONICAL_GAP.authorityRouting,
        ) &&
        !canonicalProjection
      ) {
        manager.destroy();
        mountNode.remove();
        continue;
      }
      const view = (deps.createView ?? createEditorView)(mountNode, {
        ...canonical.props,
        editable: () => !(deps.getReadOnly?.() ?? false),
        state: canonicalProjection
          ? EditorState.create({
              doc: canonicalProjection,
              plugins: [
                createDocumentNumberingPlugin(numbering),
                createDocumentStyleContextPlugin(styles),
                createNativeSelectionRefreshPlugin(),
              ],
            })
          : buildInitialState(headerFooter, styles, theme, numbering, manager),
        dispatchTransaction(transaction) {
          if ((deps.getReadOnly?.() ?? false) && transaction.docChanged) return;
          if (canonical.dispatch(transaction)) return;
          const nextState = view.state.apply(transaction);
          const docChanged = !nextState.doc.eq(view.state.doc);
          const noteIssue = noteReferenceTransactionIssue(transaction);
          if (noteIssue)
            deps.onSessionRefusal?.(
              noteIssue.message,
              { gap: CANONICAL_GAP.dispatch },
              noteIssue,
            );
          view.updateState(nextState);
          const mountedPart = mounted.get(part.rId);
          if (mountedPart && docChanged) {
            mountedPart.dirty = true;
          }
          deps.onTransaction?.({
            docChanged,
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
        appliedNumbering: numbering,
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
    snapshotDocument: (model) => {
      if (
        usesCanonicalSession(
          deps.getExperimentalSession?.(),
          CANONICAL_GAP.authorityRouting,
        )
      ) {
        const canonical = deps.getCanonicalApi?.()?.getCanonicalDocument();
        if (canonical) return canonical;
        throw new CanonicalSessionError({
          ground: { gap: CANONICAL_GAP.authorityRouting },
          reason: "refused",
          message: "Canonical story snapshot is unavailable.",
        });
      }
      const document = readModel(model);
      let headers = document.package.headers;
      let footers = document.package.footers;
      let headersChanged = false;
      let footersChanged = false;
      let numbering = document.package.numbering;

      for (const { dirty, kind, rId, view } of mounted.values()) {
        if (!dirty) {
          continue;
        }
        const source = kind === "header" ? headers : footers;
        const existing = source?.get(rId);
        if (!source || !existing) {
          continue;
        }
        // A list started in this story defined its instance on its paragraphs.
        const story = storyListNumbering(view.state, numbering);
        numbering = story.numbering;
        const updated: HeaderFooter = {
          ...existing,
          content: proseDocToBlocks(
            story.doc,
            existing.content,
            document.package.styles,
          ),
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
        return model;
      }
      const packageWithSnapshots = { ...document.package };
      if (numbering) {
        packageWithSnapshots.numbering = numbering;
      }
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
