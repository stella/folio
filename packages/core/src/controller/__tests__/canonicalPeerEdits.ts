/**
 * Concurrent canonical edits for collaboration properties: several peers,
 * each a canonical session over the same saved document, edit through the
 * session and the edits that reach the room are checked afterwards.
 *
 * Each peer types at its own caret in its own paragraph ("AAAA|BBBB") and
 * makes side edits elsewhere: a digit at the start of another peer's
 * paragraph, a filler "A" deleted there, run and paragraph formatting, a
 * split at a paragraph's end and the join that removes such an empty
 * paragraph again. Characters are unique per keystroke, so the final
 * document says which edit survived.
 *
 * The edits leave out what the sequencing contract cannot send or transform
 * yet. Typing, split and join intents compile to insertContent with
 * joinInline, and to splitBlock / joinBlocks carrying fields the wire form
 * lacks, so the peers commit those edits as the wire-form ops (insertText,
 * plain splitBlock and joinBlocks) through the session's journalled
 * `prepareOps`. Deletion and formatting go through intents. Left out
 * entirely: undo of anything but typing (insertContent, replaceInline and
 * setParagraphReview inverses, and setRunProps / setParagraphProps inverses
 * with preconditions), tables, numbering, comments, notes, package resources,
 * sections and suggestions.
 */

import { panic, type Result } from "better-result";
import { expect } from "bun:test";
import { EditorState, TextSelection } from "prosemirror-state";

import {
  captureDocumentOp,
  DOCUMENT_OP_SCHEMA_VERSION,
  type DocumentOp,
  INHERIT_RUN_PROPS,
  OP_STORIES,
  paragraphLength,
  paragraphLogicalText,
  type SequencedBatch,
  storyParagraphs,
  type TextPosition,
  transformPosition,
} from "@stll/docx-core/ops";
import { ZipArchive } from "@stll/docx-core/zip";

import { serializeCanonicalSave } from "../../docx/canonicalSave";
import { prepareCanonicalDocxInput } from "../../docx/canonicalSessionInput";
import { createDocx } from "../../docx/rezip";
import type { Document } from "../../types/document";
import type {
  CanonicalCommit,
  CanonicalSelection,
  CanonicalSession,
  CanonicalSessionError,
} from "../canonicalSession";

const HOME_TEXT = "AAAABBBB";
const CARET = 4;
const FILLER = "A";
export const MAX_PEERS = 4;
const PEER_CHARS = 0x400;
const FIRST_CHAR = 0x4e00;

export const LOCAL_KINDS = [
  "type",
  "type",
  "type",
  "type",
  "prefix",
  "deleteFiller",
  "bold",
  "align",
  "splitEnd",
  "join",
  "undo",
  "undo",
] as const;
export type LocalKind = (typeof LOCAL_KINDS)[number];

/** `rebased`: a remote batch arrived after the step was taken. */
export type HistoryMirror = { rebased: boolean } & (
  { kind: "typing"; chars: string[] } | { kind: "other" }
);
export type BatchMeta =
  | { peer: number; kind: "typing"; chars: readonly string[] }
  | { peer: number; kind: "undo"; chars: readonly string[] }
  | { peer: number; kind: "other" };

export type EditingPeer = {
  index: number;
  /** The paragraph this peer types in. */
  home: string;
  session: CanonicalSession;
  state: EditorState;
  /** The session's undo stack as this peer's own edits built it, bottom first. */
  history: HistoryMirror[];
  typed: number;
};

/** Publish `prepared` as the peer's commit described by `meta`, or skip it. */
export type CommitLocal<P extends EditingPeer> = (
  peer: P,
  prepared: Result<CanonicalCommit, CanonicalSessionError>,
  meta: BatchMeta,
) => CanonicalCommit | undefined;

let seedBytes: Uint8Array | undefined;
/** A saved DOCX with one "AAAABBBB" paragraph per possible peer. */
export const peerSeed = async (): Promise<Uint8Array> => {
  seedBytes ??= (
    await prepareCanonicalDocxInput(
      await createDocx({
        package: {
          document: {
            content: Array.from({ length: MAX_PEERS }, () => ({
              type: "paragraph" as const,
              content: [
                {
                  type: "run" as const,
                  content: [{ type: "text" as const, text: HOME_TEXT }],
                },
              ],
            })),
          },
        },
      }),
    )
  ).unwrap();
  return seedBytes;
};

const ownerOf = (char: string): number | undefined => {
  if (/^[0-9]$/u.test(char)) return Number(char);
  const code = char.codePointAt(0) ?? 0;
  if (code < FIRST_CHAR || code >= FIRST_CHAR + MAX_PEERS * PEER_CHARS)
    return undefined;
  return Math.floor((code - FIRST_CHAR) / PEER_CHARS);
};
const typedBy = (char: string, peer: number): boolean => {
  const code = char.codePointAt(0) ?? 0;
  return (
    code >= FIRST_CHAR + peer * PEER_CHARS &&
    code < FIRST_CHAR + (peer + 1) * PEER_CHARS
  );
};

export const paragraphs = (document: Document) =>
  storyParagraphs(document.package.document).map(({ paragraph }) => paragraph);
export const documentText = (document: Document): string =>
  paragraphs(document)
    .map((paragraph) => paragraphLogicalText(paragraph))
    .join("\n");
const paragraphOf = (document: Document, blockId: string) =>
  paragraphs(document).find(({ paraId }) => paraId === blockId);
export const at = (blockId: string, offset: number): TextPosition => ({
  story: OP_STORIES.MAIN,
  blockId,
  offset,
});

/** Characters other actors put in the document, counted. */
const foreignCounts = (
  document: Document,
  peer: number,
): Map<string, number> => {
  const counts = new Map<string, number>();
  for (const char of documentText(document)) {
    if (char === "\n" || typedBy(char, peer) || char === String(peer)) continue;
    counts.set(char, (counts.get(char) ?? 0) + 1);
  }
  return counts;
};

/** Every part's bytes; the save time is the one part of a save not its content. */
const partsOf = async (buffer: ArrayBuffer) => {
  const archive = await ZipArchive.read(buffer);
  const parts = new Map<string, string>();
  for (const entry of archive.entries())
    parts.set(
      entry.path,
      entry.path === "docProps/core.xml"
        ? (await entry.text()).replace(
            /<dcterms:modified\b[^>]*>[^<]*<\/dcterms:modified>/u,
            "",
          )
        : Buffer.from(await entry.bytes()).toString("hex"),
    );
  return parts;
};

/**
 * Every peer saves the same bytes, part by part, in a full save: a selective
 * save splices the blocks a session never touched from the source, and which
 * blocks a session touched depends on the order it saw the edits in.
 */
export const expectIdenticalSaves = async (
  sessions: readonly CanonicalSession[],
): Promise<void> => {
  const saves: Map<string, string>[] = [];
  for (const session of sessions) {
    const saved = await serializeCanonicalSave({
      snapshot: session.captureSaveSnapshot(),
      featureFlags: { selectiveSave: false },
    });
    saves.push(await partsOf(saved.buffer));
  }
  for (const parts of saves.slice(1)) expect(parts).toEqual(saves[0]);
};

export const caretOf = (peer: EditingPeer) =>
  peer.session.projection.selectionAt(peer.state).unwrap();
const placeCaret = (peer: EditingPeer, position: TextPosition) => {
  const shown = peer.session.projection.positionAt(position).unwrap();
  peer.state = peer.state.apply(
    peer.state.tr.setSelection(TextSelection.create(peer.state.doc, shown)),
  );
};
/**
 * The caret an edit leaves in place, moved over the edit's own operations as
 * it moves over anyone's: an insertion or a deletion ahead of it in its
 * paragraph shifts it. A caret put back where it stood before the edit would
 * land inside what the edit moved.
 */
const keptCaret = (
  caret: CanonicalSelection,
  ops: readonly DocumentOp[],
): CanonicalSelection => {
  const edit: SequencedBatch = {
    schema: DOCUMENT_OP_SCHEMA_VERSION,
    opId: "kept-caret",
    actor: "kept-caret",
    baseRev: 0,
    revision: 1,
    ops: ops.map((op) => captureDocumentOp(op)),
  };
  const move = (position: TextPosition) =>
    transformPosition(position, [edit]) ?? position;
  return { ...caret, anchor: move(caret.anchor), head: move(caret.head) };
};
/** An editor state showing `session`, its caret in the peer's own paragraph. */
export const homeState = (session: CanonicalSession, home: string) => {
  const doc = session.projection.doc;
  return EditorState.create({
    doc,
    selection: TextSelection.create(
      doc,
      session.projection.positionAt(at(home, CARET)).unwrap(),
    ),
  });
};

/** One local edit of `kind` by `peer`; `choice` picks its target. */
export const localEdit = <P extends EditingPeer>({
  peer,
  peers,
  kind,
  choice,
  commit,
  onUndoAfterRemote,
}: {
  peer: P;
  peers: readonly P[];
  kind: LocalKind;
  choice: number;
  commit: CommitLocal<P>;
  onUndoAfterRemote: () => void;
}): void => {
  const document = peer.session.document;
  const others = peers.filter(({ index }) => index !== peer.index);
  const other = others.at(choice % Math.max(1, others.length)) ?? peer;
  const caret = caretOf(peer);
  // An edit away from the caret keeps it, moved over the edit itself.
  const keeping = (ops: readonly DocumentOp[]) =>
    peer.session.prepareOps(peer.state, ops, keptCaret(caret, ops));
  const restore = (committed: CanonicalCommit) =>
    placeCaret(peer, keptCaret(caret, committed.ops).head);
  switch (kind) {
    case "type": {
      const char = String.fromCodePoint(
        FIRST_CHAR + peer.index * PEER_CHARS + peer.typed++,
      );
      const next = { ...caret.head, offset: caret.head.offset + char.length };
      const committed = commit(
        peer,
        peer.session.prepareOps(
          peer.state,
          [
            {
              type: "insertText",
              at: caret.head,
              text: char,
              runProps: INHERIT_RUN_PROPS,
            },
          ],
          { type: "text", anchor: next, head: next },
        ),
        { peer: peer.index, kind: "typing", chars: [char] },
      );
      if (committed === undefined) return;
      const top = peer.history.at(-1);
      if (committed.history === "joined" && top?.kind === "typing")
        top.chars.push(char);
      else peer.history.push({ kind: "typing", chars: [char], rebased: false });
      return;
    }
    case "prefix": {
      const committed = commit(
        peer,
        keeping([
          {
            type: "insertText",
            at: at(other.home, 0),
            text: String(peer.index),
            runProps: INHERIT_RUN_PROPS,
          },
        ]),
        { peer: peer.index, kind: "other" },
      );
      if (committed === undefined) return;
      peer.history.push({ kind: "other", rebased: false });
      return;
    }
    case "deleteFiller": {
      const target = paragraphOf(document, other.home);
      if (target === undefined) return;
      if (paragraphLogicalText(target).at(0) !== FILLER) return;
      const committed = commit(
        peer,
        peer.session.prepareIntent(peer.state, {
          type: "replaceText",
          from: at(other.home, 0),
          to: at(other.home, 1),
          text: "",
        }),
        { peer: peer.index, kind: "other" },
      );
      if (committed === undefined) return;
      peer.history.push({ kind: "other", rebased: false });
      restore(committed);
      return;
    }
    case "bold":
    case "align": {
      const owner = choice % 2 === 0 ? peer : other;
      const target = paragraphOf(document, owner.home);
      if (target === undefined) return;
      const committed = commit(
        peer,
        peer.session.prepareIntent(
          peer.state,
          kind === "bold"
            ? {
                type: "formatRun",
                from: at(owner.home, 0),
                to: at(owner.home, paragraphLength(target)),
                patch: { bold: choice % 3 !== 0 },
              }
            : {
                type: "formatParagraph",
                at: at(owner.home, 0),
                patch: { alignment: choice % 3 === 0 ? "center" : "right" },
              },
        ),
        { peer: peer.index, kind: "other" },
      );
      if (committed === undefined) return;
      peer.history.push({ kind: "other", rebased: false });
      restore(committed);
      return;
    }
    case "splitEnd": {
      const owner = choice % 2 === 0 ? peer : other;
      const target = paragraphOf(document, owner.home);
      if (target === undefined) return;
      const fresh = peer.session.freshBlockIds([1]);
      // An exhausted lease mints nothing until the authority renews it.
      if (fresh.isErr()) return;
      const newBlockId = fresh.value[0]?.[0];
      if (newBlockId === undefined) return;
      const committed = commit(
        peer,
        keeping([
          {
            type: "splitBlock",
            at: at(owner.home, paragraphLength(target)),
            newBlockId,
            newHalf: "second",
          },
        ]),
        { peer: peer.index, kind: "other" },
      );
      if (committed === undefined) return;
      peer.history.push({ kind: "other", rebased: false });
      return;
    }
    case "join": {
      const list = paragraphs(document);
      const empty = list.filter(
        (paragraph, index) => index > 0 && paragraphLength(paragraph) === 0,
      );
      const next = empty.at(choice % Math.max(1, empty.length));
      const previous =
        next === undefined ? undefined : list.at(list.indexOf(next) - 1);
      if (next?.paraId === undefined || previous?.paraId === undefined) return;
      const committed = commit(
        peer,
        keeping([
          {
            type: "joinBlocks",
            story: OP_STORIES.MAIN,
            blockId: previous.paraId,
            nextBlockId: next.paraId,
            survivor: "first",
          },
        ]),
        { peer: peer.index, kind: "other" },
      );
      if (committed === undefined) return;
      peer.history.push({ kind: "other", rebased: false });
      return;
    }
    case "undo": {
      const top = peer.history.at(-1);
      if (top?.kind !== "typing" || !peer.session.canUndo) return;
      const foreign = foreignCounts(document, peer.index);
      const committed = commit(peer, peer.session.prepareUndo(peer.state), {
        peer: peer.index,
        kind: "undo",
        chars: top.chars,
      });
      if (committed === undefined) return;
      peer.history.pop();
      if (top.rebased) onUndoAfterRemote();
      // Undo takes back this peer's own step, never another actor's edit.
      expect(foreignCounts(peer.session.document, peer.index)).toEqual(foreign);
      return;
    }
    default: {
      const exhaustive: never = kind;
      return exhaustive;
    }
  }
};

/**
 * No sequenced edit is lost: each peer's characters in the final text are
 * exactly the keystrokes the room sequenced, less the ones it undid. A peer
 * whose caret was never reset (`contiguous`) keeps them together, in order.
 */
export const expectSequencedKeystrokes = ({
  finalText,
  journal,
  batches,
  peers,
  contiguous,
}: {
  finalText: string;
  journal: readonly { opId: string }[];
  batches: ReadonlyMap<string, BatchMeta>;
  peers: number;
  contiguous: (peer: number) => boolean;
}): void => {
  const runs = Array.from({ length: peers }, () => [] as string[]);
  for (const { opId } of journal) {
    const meta = batches.get(opId);
    if (meta === undefined) return panic(`Unknown sequenced batch ${opId}.`);
    const run = runs.at(meta.peer) ?? [];
    if (meta.kind === "typing") run.push(...meta.chars);
    if (meta.kind === "undo")
      for (const char of meta.chars) {
        const undone = run.indexOf(char);
        expect(undone).toBeGreaterThanOrEqual(0);
        run.splice(undone, 1);
      }
  }
  for (const [peer, run] of runs.entries()) {
    const present = [...finalText].filter((char) => typedBy(char, peer));
    expect(new Set(present)).toEqual(new Set(run));
    expect(present).toHaveLength(run.length);
    if (contiguous(peer)) expect(finalText).toContain(run.join(""));
  }
  for (const char of finalText) {
    const owner = ownerOf(char);
    if (owner !== undefined) expect(owner).toBeLessThan(peers);
  }
};
