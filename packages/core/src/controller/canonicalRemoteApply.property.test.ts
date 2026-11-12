import { panic, type Result } from "better-result";
import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { EditorState, TextSelection } from "prosemirror-state";

import {
  type ClientDocumentChange,
  createClient,
  createSequencer,
  DOCUMENT_OP_SCHEMA_VERSION,
  OP_STORIES,
  paragraphLength,
  paragraphLogicalText,
  type SequencedBatch,
  storyParagraphs,
  type TextPosition,
} from "@stll/docx-core/ops";
import { ZipArchive } from "@stll/docx-core/zip";

import {
  assertProperty,
  propertyConfig,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { serializeCanonicalSave } from "../docx/canonicalSave";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import type { Document } from "../types/document";
import {
  type CanonicalCommit,
  type CanonicalSession,
  type CanonicalSessionError,
  createCanonicalSession,
  publishCanonicalProjection,
} from "./canonicalSession";

setDefaultTimeout(propertyTestTimeout(120_000));

/*
 * N peers, each a canonical session over the same saved document, edit
 * concurrently through the real sequencer and optimistic client; every
 * remote batch enters a session through `prepareRemote`.
 *
 * Each peer types at its own caret in its own paragraph ("AAAA|BBBB") and
 * makes side edits elsewhere: a digit at the start of another peer's
 * paragraph, a filler "A" deleted there, run and paragraph formatting, a
 * split at a paragraph's end and the join that removes such an empty
 * paragraph again. Characters are unique per keystroke, so the final
 * document says which edit survived.
 *
 * The generator leaves out the operations whose wire form or transform the
 * sequencing contract does not define yet: undo of anything but typing
 * (insertContent, replaceInline and setParagraphReview inverses, and
 * setRunProps / setParagraphProps inverses with preconditions), tables,
 * numbering, comments, notes, package resources, sections and suggestions.
 */

const HOME_TEXT = "AAAABBBB";
const CARET = 4;
const FILLER = "A";
const MAX_PEERS = 4;
const PEER_CHARS = 0x400;
const FIRST_CHAR = 0x4e00;

const LOCAL_KINDS = [
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
type LocalKind = (typeof LOCAL_KINDS)[number];

type Command =
  | { type: "local"; peer: number; kind: LocalKind; seed: number }
  | { type: "deliver"; seed: number }
  | { type: "submit"; peer: number };

const commandArbitrary: fc.Arbitrary<Command> = fc.oneof(
  {
    weight: 5,
    arbitrary: fc.record({
      type: fc.constant("local" as const),
      peer: fc.nat({ max: MAX_PEERS - 1 }),
      kind: fc.constantFrom(...LOCAL_KINDS),
      seed: fc.nat(),
    }),
  },
  {
    weight: 4,
    arbitrary: fc.record({
      type: fc.constant("deliver" as const),
      seed: fc.nat(),
    }),
  },
  {
    weight: 3,
    arbitrary: fc.record({
      type: fc.constant("submit" as const),
      peer: fc.nat({ max: MAX_PEERS - 1 }),
    }),
  },
);

/** How a peer applies a remote change: the session's entry, or a deliberate break. */
type RemoteStrategy = "session" | "journaledLocal" | "editorPosition";

/** `rebased`: a remote batch arrived after the step was taken. */
type HistoryMirror = { rebased: boolean } & (
  { kind: "typing"; chars: string[] } | { kind: "other" }
);
type BatchMeta =
  | { peer: number; kind: "typing"; chars: readonly string[] }
  | { peer: number; kind: "undo"; chars: readonly string[] }
  | { peer: number; kind: "other" };

type Client = ReturnType<typeof createClient>;
type Sequencer = ReturnType<typeof createSequencer>;
type Submission = ReturnType<Sequencer["submit"]>;

type Peer = {
  index: number;
  actor: string;
  home: string;
  session: CanonicalSession;
  client: Client;
  state: EditorState;
  /** The session's undo stack as this peer's own edits built it, bottom first. */
  history: HistoryMirror[];
  /** A pending batch was dropped: the caret's continuity is not promised. */
  replaced: boolean;
  typed: number;
  splits: number;
  ops: number;
};

type Delivery =
  | { peer: number; type: "result"; result: Submission }
  | { peer: number; type: "broadcast"; batch: SequencedBatch };

type Coverage = {
  remoteApplied: number;
  undoAfterRemote: number;
  droppedSteps: number;
  rejected: number;
};

let seedBytes: Uint8Array | undefined;
const seed = async (): Promise<Uint8Array> => {
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

const paragraphs = (document: Document) =>
  storyParagraphs(document.package.document).map(({ paragraph }) => paragraph);
const documentText = (document: Document): string =>
  paragraphs(document)
    .map((paragraph) => paragraphLogicalText(paragraph))
    .join("\n");
const paragraphOf = (document: Document, blockId: string) =>
  paragraphs(document).find(({ paraId }) => paraId === blockId);
const at = (blockId: string, offset: number): TextPosition => ({
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

const partsOf = async (buffer: ArrayBuffer) => {
  const archive = await ZipArchive.read(buffer);
  const parts = new Map<string, string>();
  for (const entry of archive.entries())
    parts.set(entry.path, Buffer.from(await entry.bytes()).toString("hex"));
  return parts;
};

const simulate = async (
  peersCount: number,
  commands: readonly Command[],
  strategy: RemoteStrategy,
  coverage: Coverage,
): Promise<void> => {
  const bytes = await seed();
  const peers: Peer[] = [];
  let sequencer: Sequencer | undefined;
  for (let index = 0; index < peersCount; index += 1) {
    const parsed = await parseDocx(bytes, {
      preloadFonts: false,
      detectVariables: false,
    });
    const session = createCanonicalSession(parsed).unwrap();
    const home = paragraphs(session.document).at(index)?.paraId;
    if (home === undefined) return panic("The seed has a paragraph per peer.");
    sequencer ??= createSequencer(session.document);
    const doc = session.projection.doc;
    const peer: Peer = {
      index,
      actor: `actor-${index}`,
      home,
      session,
      client: createClient(session.document),
      state: EditorState.create({
        doc,
        selection: TextSelection.create(
          doc,
          session.projection.positionAt(at(home, CARET)).unwrap(),
        ),
      }),
      history: [],
      replaced: false,
      typed: 0,
      splits: 0,
      ops: 0,
    };
    // The host mirror drops what the session drops, from the bottom.
    session.onHistoryDrop(({ undo }) => {
      coverage.droppedSteps += undo;
      peer.history.splice(0, undo);
    });
    peers.push(peer);
  }
  if (sequencer === undefined) return panic("A room has a peer.");
  const room = sequencer;
  const batches = new Map<string, BatchMeta>();
  const deliveries: Delivery[] = [];

  const caretOf = (peer: Peer) =>
    peer.session.projection.selectionAt(peer.state).unwrap();
  const placeCaret = (peer: Peer, position: TextPosition) => {
    const shown = peer.session.projection.positionAt(position).unwrap();
    peer.state = peer.state.apply(
      peer.state.tr.setSelection(TextSelection.create(peer.state.doc, shown)),
    );
  };

  /** Enqueue first: a commit the room cannot take is never published. */
  const commitLocal = (
    peer: Peer,
    prepared: Result<CanonicalCommit, CanonicalSessionError>,
    meta: BatchMeta,
  ): CanonicalCommit | undefined => {
    if (prepared.isErr()) return undefined;
    const commit = prepared.value;
    const opId = `${peer.actor}-${peer.ops++}`;
    peer.client
      .enqueue({
        schema: DOCUMENT_OP_SCHEMA_VERSION,
        opId,
        actor: peer.actor,
        baseRev: peer.client.headRev,
        ops: commit.ops,
      })
      .unwrap();
    peer.state = publishCanonicalProjection({
      state: peer.state,
      commit,
      session: peer.session,
    }).unwrap().state;
    batches.set(opId, meta);
    return commit;
  };

  const applyRemote = (peer: Peer, change: ClientDocumentChange) => {
    if (change.type === "replaced") peer.replaced = true;
    coverage.remoteApplied += 1;
    if (strategy === "journaledLocal") {
      const prepared = peer.session.prepareOps(
        peer.state,
        change.batch.ops,
        caretOf(peer),
      );
      if (prepared.isErr() && prepared.error.reason === "noChange") return;
      // The host never learns that the session took the remote batch as
      // its own step.
      peer.state = publishCanonicalProjection({
        state: peer.state,
        commit: prepared.unwrap(),
        session: peer.session,
      }).unwrap().state;
      return;
    }
    for (const step of peer.history) step.rebased = true;
    const before = peer.state.selection;
    const prepared = peer.session.prepareRemote(peer.state, change);
    if (prepared.isErr() && prepared.error.reason === "noChange") return;
    const commit = prepared.unwrap();
    peer.state = publishCanonicalProjection({
      state: peer.state,
      commit,
      session: peer.session,
    }).unwrap().state;
    if (strategy === "editorPosition")
      peer.state = peer.state.apply(
        peer.state.tr.setSelection(
          before.map(peer.state.doc, commit.transaction.mapping),
        ),
      );
  };

  const local = (peer: Peer, kind: LocalKind, choice: number) => {
    const document = peer.session.document;
    const others = peers.filter(({ index }) => index !== peer.index);
    const other = others.at(choice % Math.max(1, others.length)) ?? peer;
    const caret = caretOf(peer);
    const restore = () => placeCaret(peer, caret.head);
    switch (kind) {
      case "type": {
        const char = String.fromCodePoint(
          FIRST_CHAR + peer.index * PEER_CHARS + peer.typed++,
        );
        const head = peer.state.selection.head;
        const commit = commitLocal(
          peer,
          peer.session.prepareReplace(peer.state, {
            from: head,
            to: head,
            text: char,
            semantic: "typing",
            time: peer.typed * 1000,
          }),
          { peer: peer.index, kind: "typing", chars: [char] },
        );
        if (commit === undefined) return;
        const top = peer.history.at(-1);
        if (commit.history === "joined" && top?.kind === "typing")
          top.chars.push(char);
        else
          peer.history.push({ kind: "typing", chars: [char], rebased: false });
        return;
      }
      case "prefix":
      case "deleteFiller": {
        const target = paragraphOf(document, other.home);
        if (target === undefined) return;
        if (
          kind === "deleteFiller" &&
          paragraphLogicalText(target).at(0) !== FILLER
        )
          return;
        const commit = commitLocal(
          peer,
          peer.session.prepareIntent(peer.state, {
            type: "replaceText",
            from: at(other.home, 0),
            to: at(other.home, kind === "prefix" ? 0 : 1),
            text: kind === "prefix" ? String(peer.index) : "",
          }),
          { peer: peer.index, kind: "other" },
        );
        if (commit === undefined) return;
        peer.history.push({ kind: "other", rebased: false });
        restore();
        return;
      }
      case "bold":
      case "align": {
        const owner = choice % 2 === 0 ? peer : other;
        const target = paragraphOf(document, owner.home);
        if (target === undefined) return;
        const commit = commitLocal(
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
        if (commit === undefined) return;
        peer.history.push({ kind: "other", rebased: false });
        restore();
        return;
      }
      case "splitEnd": {
        const owner = choice % 2 === 0 ? peer : other;
        const target = paragraphOf(document, owner.home);
        if (target === undefined) return;
        // A distinct id block per actor stands in for the authority's leases.
        const newBlockId = (0x70000000 + peer.index * 0x100000 + peer.splits++)
          .toString(16)
          .toUpperCase();
        const commit = commitLocal(
          peer,
          peer.session.prepareIntent(peer.state, {
            type: "splitParagraph",
            at: at(owner.home, paragraphLength(target)),
            newBlockId,
            newHalf: "second",
          }),
          { peer: peer.index, kind: "other" },
        );
        if (commit === undefined) return;
        peer.history.push({ kind: "other", rebased: false });
        restore();
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
        if (next?.paraId === undefined || previous?.paraId === undefined)
          return;
        const commit = commitLocal(
          peer,
          peer.session.prepareIntent(peer.state, {
            type: "joinParagraphs",
            story: OP_STORIES.MAIN,
            blockId: previous.paraId,
            nextBlockId: next.paraId,
            survivor: "first",
          }),
          { peer: peer.index, kind: "other" },
        );
        if (commit === undefined) return;
        peer.history.push({ kind: "other", rebased: false });
        restore();
        return;
      }
      case "undo": {
        const top = peer.history.at(-1);
        if (top?.kind !== "typing" || !peer.session.canUndo) return;
        const foreign = foreignCounts(document, peer.index);
        const commit = commitLocal(peer, peer.session.prepareUndo(peer.state), {
          peer: peer.index,
          kind: "undo",
          chars: top.chars,
        });
        if (commit === undefined) return;
        peer.history.pop();
        if (top.rebased) coverage.undoAfterRemote += 1;
        // Undo takes back this peer's own step, never another actor's edit.
        expect(foreignCounts(peer.session.document, peer.index)).toEqual(
          foreign,
        );
        return;
      }
      default: {
        const exhaustive: never = kind;
        return exhaustive;
      }
    }
  };

  const submit = (peer: Peer) => {
    const batch = peer.client.nextSubmission();
    if (batch === undefined) return;
    const result = room.submit(batch);
    deliveries.push({ peer: peer.index, type: "result", result });
    if (result.type === "reject") {
      coverage.rejected += 1;
      return;
    }
    const broadcast = room.broadcasts.at(-1);
    if (broadcast === undefined) return;
    for (const target of peers)
      deliveries.push({
        peer: target.index,
        type: "broadcast",
        batch: broadcast,
      });
  };

  const deliver = (index: number) => {
    const event = deliveries.splice(index % deliveries.length, 1).at(0);
    if (event === undefined) return;
    const peer = peers.at(event.peer);
    if (peer === undefined) return;
    if (event.type === "broadcast") {
      for (const change of peer.client.receiveBroadcast(event.batch).unwrap())
        applyRemote(peer, change);
    } else if (event.result.type === "ack")
      peer.client.receiveAck(event.result);
    else
      for (const change of peer.client.receiveReject(event.result))
        applyRemote(peer, change);
  };

  for (const command of commands) {
    switch (command.type) {
      case "local": {
        const peer = peers.at(command.peer % peersCount);
        if (peer !== undefined) local(peer, command.kind, command.seed);
        break;
      }
      case "submit": {
        const peer = peers.at(command.peer % peersCount);
        if (peer !== undefined) submit(peer);
        break;
      }
      case "deliver":
        if (deliveries.length > 0) deliver(command.seed);
        break;
      default: {
        const exhaustive: never = command;
        return exhaustive;
      }
    }
  }
  for (let steps = 0; ; steps += 1) {
    expect(steps).toBeLessThan(10_000);
    for (const peer of peers) submit(peer);
    if (deliveries.length === 0) break;
    deliver(0);
  }

  const final = room.document;
  const finalText = documentText(final);
  const saves: Map<string, string>[] = [];
  for (const peer of peers) {
    expect(peer.client.pending).toHaveLength(0);
    expect(peer.client.headRev).toBe(room.headRev);
    expect(documentText(peer.session.document)).toBe(finalText);
    const saved = await serializeCanonicalSave({
      snapshot: peer.session.captureSaveSnapshot(),
    });
    saves.push(await partsOf(saved.buffer));
  }
  // Every peer saves the same bytes, part by part.
  for (const parts of saves.slice(1)) expect(parts).toEqual(saves[0]);

  // No acknowledged edit is lost: a peer's characters in the final document
  // are exactly the keystrokes the room sequenced, less the ones it undid.
  const runs = peers.map(() => [] as string[]);
  for (const { opId } of room.broadcasts) {
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
  for (const peer of peers) {
    const run = runs.at(peer.index) ?? [];
    const present = [...finalText].filter((char) => typedBy(char, peer.index));
    expect(new Set(present)).toEqual(new Set(run));
    expect(present).toHaveLength(run.length);
    // The caret followed the peer's own typing across every remote edit:
    // its keystrokes stand together, in order.
    if (!peer.replaced) expect(finalText).toContain(run.join(""));
    for (const char of finalText) {
      const owner = ownerOf(char);
      if (owner !== undefined) expect(owner).toBeLessThan(peersCount);
    }
  }
};

const law = (strategy: RemoteStrategy, coverage: Coverage) =>
  fc.asyncProperty(
    fc.integer({ min: 2, max: MAX_PEERS }),
    fc.array(commandArbitrary, { minLength: 10, maxLength: 60 }),
    async (peersCount, commands) =>
      simulate(peersCount, commands, strategy, coverage),
  );

const emptyCoverage = (): Coverage => ({
  remoteApplied: 0,
  undoAfterRemote: 0,
  droppedSteps: 0,
  rejected: 0,
});

test("peers applying remote batches through the session converge and keep their own history", async () => {
  const coverage = emptyCoverage();
  await assertProperty(law("session", coverage), {
    numRuns: 40,
    id: "canonical remote apply convergence",
  });
  expect(coverage.remoteApplied).toBeGreaterThan(0);
  expect(coverage.undoAfterRemote).toBeGreaterThan(0);
});

test("the property catches a remote batch journaled as a local undo step", async () => {
  const result = await fc.check(
    law("journaledLocal", emptyCoverage()),
    propertyConfig({ numRuns: 100, endOnFailure: true }),
  );
  expect(result.failed).toBe(true);
});

test("the property catches a selection mapped by editor position", async () => {
  const result = await fc.check(
    law("editorPosition", emptyCoverage()),
    propertyConfig({ numRuns: 100, endOnFailure: true }),
  );
  expect(result.failed).toBe(true);
});
