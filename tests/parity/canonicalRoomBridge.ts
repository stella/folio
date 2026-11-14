import type { FolioEditor } from "@stll/folio-core/controller/folioEditor";

import { withoutMintedIds } from "./canonicalBridge";

/** A playground's place in its canonical room, as the room bridge reports it. */
export type CanonicalRoomBridgeStatus = {
  /** The editor's transport reaches the room's authority. */
  connected: boolean;
  /** This tab runs the room's authority. */
  leading: boolean;
  /** The other peers whose presence this tab holds, by actor. */
  peers: readonly string[];
  /** The last room event that is not routine: a refusal, a stop. */
  notice: string | undefined;
};

/**
 * Private interaction-test bridge for the canonical room both playgrounds
 * run with `?collaboration=1`. Tests type with the keyboard; comments go
 * through the editor's canonical comment API, as its sidebar sends them.
 */
export const buildCanonicalRoomBridge = (
  getEditor: () => FolioEditor | null,
  getStatus: () => CanonicalRoomBridgeStatus,
) => {
  const createComment = (
    anchor:
      | { kind: "selection"; from: number; to: number }
      | { kind: "reply"; parentId: number },
    text: string,
  ): number | null => {
    const result = getEditor()?.applyCanonicalComment({
      type: "create",
      text,
      author: "Room tester",
      anchor: anchor.kind === "reply" ? anchor : { ...anchor, story: "main" },
    });
    return result?.status === "applied" ? (result.commentId ?? null) : null;
  };
  return {
    status: getStatus,
    /** Whether the editor takes local edits: it joined the room. */
    editable: () => getEditor()?.getView()?.editable ?? false,
    /** The editor position where the body's `index`-th paragraph starts. */
    paragraphStart: (index: number): number | null => {
      const view = getEditor()?.getView();
      if (!view) return null;
      let seen = 0;
      let found: number | null = null;
      view.state.doc.descendants((node, position) => {
        if (found !== null) return false;
        if (!node.isTextblock) return true;
        if (seen === index) found = position + 1;
        seen += 1;
        return false;
      });
      return found;
    },
    /**
     * The session's body, comments and notes as JSON: the same string in
     * every tab that applied the same journal.
     */
    document: (): string | null => {
      const document = getEditor()?.getCanonicalDocument();
      if (!document) return null;
      const { content, comments = [] } = document.package.document;
      const { footnotes = [], endnotes = [] } = document.package;
      return JSON.stringify(
        withoutMintedIds({ content, comments, footnotes, endnotes }),
      );
    },
    comments: () =>
      getEditor()
        ?.getCanonicalComments()
        ?.map(({ id, parentId, author, done }) => ({
          id,
          parentId: parentId ?? null,
          author,
          done: done ?? false,
        })) ?? null,
    comment: (from: number, to: number, text: string) =>
      createComment({ kind: "selection", from, to }, text),
    reply: (parentId: number, text: string) =>
      createComment({ kind: "reply", parentId }, text),
  };
};

export type CanonicalRoomBridge = ReturnType<typeof buildCanonicalRoomBridge>;

declare global {
  var __folioCanonicalRoom: CanonicalRoomBridge | undefined;
}
