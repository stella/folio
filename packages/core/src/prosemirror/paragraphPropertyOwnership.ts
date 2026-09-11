import { panic } from "better-result";
import type { Node as PMNode } from "prosemirror-model";
import type { Transaction } from "prosemirror-state";
import { Step, StepMap, StepResult, type Mappable } from "prosemirror-transform";
import { ySyncPluginKey } from "y-prosemirror";

import {
  PARAGRAPH_PROPERTY_SOURCE_MAX_PARAGRAPHS,
  ParagraphPropertySourceContract,
  ParagraphPropertySourceValidationError,
  type ParagraphPropertySourceStory,
} from "../docx/paragraphPropertySourceIdentity";
import { PROSE_PARAGRAPH_SOURCE_CONTRACT_ATTR } from "../docx/paragraphPropertySource";
import { readParagraphPropertyState, type ParagraphPropertyState } from "./paragraphPropertyState";

type ImportedOccurrence = {
  node: PMNode;
  nodeSize: number;
  pos: number;
  token: string;
};

const importedOccurrenceMatches = (left: ImportedOccurrence, right: ImportedOccurrence): boolean =>
  left.token === right.token && left.pos === right.pos && left.nodeSize === right.nodeSize;

type ParagraphOccurrence = {
  node: PMNode;
  pos: number;
  state: ParagraphPropertyState;
};

type OwnershipSnapshot = {
  imported: ReadonlyMap<string, ImportedOccurrence>;
  paragraphs: readonly ParagraphOccurrence[];
  paragraphCount: number;
};

const EDITABLE_DOCUMENT_STORY = Object.freeze({
  type: "document",
}) satisfies ParagraphPropertySourceStory;

const readContract = (doc: PMNode): ParagraphPropertySourceContract | null => {
  const result = ParagraphPropertySourceContract.read(
    doc.attrs[PROSE_PARAGRAPH_SOURCE_CONTRACT_ATTR],
  );
  if (result.status === "invalid") {
    throw new ParagraphPropertySourceValidationError({
      code: "contract_mismatch",
      message: "The ProseMirror document contains an invalid paragraph-property source contract.",
    });
  }
  return result.status === "valid" ? result.value : null;
};

const ownershipSnapshot = (
  doc: PMNode,
  contract: ParagraphPropertySourceContract | null,
  story: ParagraphPropertySourceStory,
): OwnershipSnapshot => {
  const imported = new Map<string, ImportedOccurrence>();
  const paragraphs: ParagraphOccurrence[] = [];
  let paragraphCount = 0;
  doc.descendants((node, pos) => {
    if (node.type.name !== "paragraph") {
      return true;
    }
    paragraphCount += 1;
    if (paragraphCount > PARAGRAPH_PROPERTY_SOURCE_MAX_PARAGRAPHS) {
      throw new ParagraphPropertySourceValidationError({
        code: "source_capacity_exceeded",
        message: "Paragraph-property transaction census capacity was exceeded.",
      });
    }
    const state = readParagraphPropertyState(node.attrs["_paragraphPropertyState"]);
    if (state.status !== "valid") {
      throw new ParagraphPropertySourceValidationError({
        code: "invalid_state",
        message: "A paragraph contains invalid paragraph-property state.",
        ...(state.status === "invalid" ? { token: state.raw } : {}),
      });
    }
    paragraphs.push(Object.freeze({ node, pos, state: state.value }));
    if (state.value.type !== "imported") {
      return false;
    }
    if (!contract) {
      throw new ParagraphPropertySourceValidationError({
        code: "contract_mismatch",
        message: "An imported paragraph-property state requires its source contract.",
        token: state.value.token,
      });
    }
    const token = contract.readToken(state.value.token);
    if (token.status !== "valid" || !token.value.belongsToStory(story)) {
      throw new ParagraphPropertySourceValidationError({
        code: "invalid_token",
        message: "A paragraph-property token is outside the document ownership context.",
        token: state.value.token,
      });
    }
    if (imported.has(state.value.token)) {
      throw new ParagraphPropertySourceValidationError({
        code: "duplicate_token",
        message: "A paragraph-property token is attached to more than one paragraph.",
        token: state.value.token,
      });
    }
    imported.set(
      state.value.token,
      Object.freeze({ node, nodeSize: node.nodeSize, pos, token: state.value.token }),
    );
    return false;
  });
  return {
    imported,
    paragraphs: Object.freeze(paragraphs),
    paragraphCount,
  };
};

type ParagraphPropertyOwnershipTransition =
  | {
      retained: ImportedOccurrence | null;
      type: "split-left-created-right-retains";
    }
  | {
      displaced: ImportedOccurrence | null;
      retained: ImportedOccurrence | null;
      sourceFrom: number;
      sourceTo: number;
      type: "join-left-paragraph-mark-retains" | "join-right-paragraph-mark-retains";
    }
  | {
      affected: readonly ImportedOccurrence[];
      type: "delete-selection";
    }
  | {
      affected: readonly ImportedOccurrence[];
      retained: ImportedOccurrence | null;
      type: "replace-selection-then-split-right-retains";
    };

const paragraphPropertyOwnershipStepIssuer = Symbol("paragraphPropertyOwnershipStepIssuer");
type ParagraphPropertyOwnershipStepIssuer = typeof paragraphPropertyOwnershipStepIssuer;

class ParagraphPropertyOwnershipTransitionStep extends Step {
  readonly #direction: "forward" | "inverse";
  readonly #transition: ParagraphPropertyOwnershipTransition;

  constructor(
    issuer: ParagraphPropertyOwnershipStepIssuer,
    transition: ParagraphPropertyOwnershipTransition,
    direction: "forward" | "inverse" = "forward",
  ) {
    super();
    if (issuer !== paragraphPropertyOwnershipStepIssuer) {
      panic("Only paragraph-property structural sinks may issue ownership steps");
    }
    this.#transition = Object.freeze(transition);
    this.#direction = direction;
    Object.freeze(this);
  }

  apply(doc: PMNode): StepResult {
    const snapshot = ownershipSnapshot(doc, readContract(doc), EDITABLE_DOCUMENT_STORY);
    const expected: ImportedOccurrence[] = [];
    switch (this.#transition.type) {
      case "delete-selection":
        expected.push(...this.#transition.affected);
        break;
      case "replace-selection-then-split-right-retains":
        expected.push(...this.#transition.affected);
        break;
      case "join-left-paragraph-mark-retains":
      case "join-right-paragraph-mark-retains":
        if (this.#transition.retained !== null) {
          expected.push(this.#transition.retained);
        }
        if (this.#transition.displaced !== null) {
          expected.push(this.#transition.displaced);
        }
        break;
      case "split-left-created-right-retains":
        if (this.#transition.retained !== null) {
          expected.push(this.#transition.retained);
        }
        break;
      default: {
        const exhaustive: never = this.#transition;
        return exhaustive;
      }
    }
    for (const occurrence of expected) {
      const current = snapshot.imported.get(occurrence.token);
      if (!current || !importedOccurrenceMatches(current, occurrence)) {
        return StepResult.fail(
          "Paragraph-property ownership proof does not match the transaction source census",
        );
      }
    }
    return StepResult.ok(doc);
  }

  getMap(): StepMap {
    return StepMap.empty;
  }

  invert(): Step {
    return new ParagraphPropertyOwnershipTransitionStep(
      paragraphPropertyOwnershipStepIssuer,
      this.#transition,
      this.#direction === "forward" ? "inverse" : "forward",
    );
  }

  map(_mapping: Mappable): Step {
    return this;
  }

  merge(_other: Step): null {
    return null;
  }

  toJSON(): never {
    return panic("Paragraph-property ownership proof steps cannot cross a wire boundary");
  }

  read(): {
    direction: "forward" | "inverse";
    transition: ParagraphPropertyOwnershipTransition;
  } {
    return { direction: this.#direction, transition: this.#transition };
  }
}

const importedOccurrence = (
  state: ParagraphPropertyState,
  node: PMNode,
  pos: number,
): ImportedOccurrence | null =>
  state.type === "imported"
    ? Object.freeze({ node, nodeSize: node.nodeSize, pos, token: state.token })
    : null;

const importedOccurrenceAt = (node: PMNode, pos: number): ImportedOccurrence | null => {
  if (node.type.name !== "paragraph") {
    panic("Paragraph-property ownership can only describe paragraphs");
  }
  const state = readParagraphPropertyState(node.attrs["_paragraphPropertyState"]);
  if (state.status !== "valid") {
    throw new ParagraphPropertySourceValidationError({
      code: "invalid_state",
      message: "A paragraph-property ownership sink received invalid state.",
      ...(state.status === "invalid" ? { token: state.raw } : {}),
    });
  }
  return importedOccurrence(state.value, node, pos);
};

const appendProof = (
  transaction: Transaction,
  transition: ParagraphPropertyOwnershipTransition,
): void => {
  transaction.step(
    new ParagraphPropertyOwnershipTransitionStep(paragraphPropertyOwnershipStepIssuer, transition),
  );
};

type RecordSplitOwnershipProofOptions = {
  pos: number;
  transaction: Transaction;
};

export const recordSplitParagraphPropertyOwnershipProof = ({
  pos,
  transaction,
}: RecordSplitOwnershipProofOptions): void => {
  const $pos = transaction.doc.resolve(pos);
  const retained = importedOccurrenceAt($pos.parent, $pos.before());
  if (retained === null) {
    return;
  }
  appendProof(transaction, {
    type: "split-left-created-right-retains",
    retained,
  });
};

type RecordJoinOwnershipProofOptions = {
  joinPos: number;
  transaction: Transaction;
  transition: {
    type: "join-left-paragraph-mark-retains" | "join-right-paragraph-mark-retains";
  };
};

export const recordJoinParagraphPropertyOwnershipProof = ({
  joinPos,
  transaction,
  transition,
}: RecordJoinOwnershipProofOptions): void => {
  const $join = transaction.doc.resolve(joinPos);
  const left = $join.nodeBefore;
  const right = $join.nodeAfter;
  if (left?.type.name !== "paragraph" || right?.type.name !== "paragraph") {
    panic("Paragraph-property join proof requires adjacent paragraphs");
  }
  const leftOccurrence = importedOccurrenceAt(left, joinPos - left.nodeSize);
  const rightOccurrence = importedOccurrenceAt(right, joinPos);
  if (leftOccurrence === null && rightOccurrence === null) {
    return;
  }
  appendProof(transaction, {
    type: transition.type,
    retained:
      transition.type === "join-left-paragraph-mark-retains" ? leftOccurrence : rightOccurrence,
    displaced:
      transition.type === "join-left-paragraph-mark-retains" ? rightOccurrence : leftOccurrence,
    sourceFrom: joinPos - left.nodeSize,
    sourceTo: joinPos + right.nodeSize,
  });
};

export const recordDeleteParagraphPropertyOwnershipProof = (transaction: Transaction): void => {
  if (transaction.selection.empty) {
    panic("A paragraph deletion proof requires a non-empty selection");
  }
  const affected: ImportedOccurrence[] = [];
  transaction.doc.nodesBetween(
    transaction.selection.from,
    transaction.selection.to,
    (node, pos) => {
      if (node.type.name !== "paragraph") {
        return true;
      }
      const occurrence = importedOccurrenceAt(node, pos);
      if (occurrence) {
        affected.push(occurrence);
      }
      return false;
    },
  );
  if (affected.length === 0) {
    return;
  }
  appendProof(transaction, {
    type: "delete-selection",
    affected: Object.freeze(affected),
  });
};

export const recordReplaceSelectionThenSplitParagraphPropertyOwnershipProof = ({
  transaction,
}: {
  transaction: Transaction;
}): void => {
  if (transaction.selection.empty || transaction.selection.$from.parent.type.name !== "paragraph") {
    panic("A replace-and-split proof requires a paragraph selection");
  }
  const retained = importedOccurrenceAt(
    transaction.selection.$from.parent,
    transaction.selection.$from.before(),
  );
  const affected: ImportedOccurrence[] = [];
  transaction.doc.nodesBetween(
    transaction.selection.from,
    transaction.selection.to,
    (node, pos) => {
      if (node.type.name !== "paragraph") {
        return true;
      }
      const occurrence = importedOccurrenceAt(node, pos);
      if (occurrence) {
        affected.push(occurrence);
      }
      return false;
    },
  );
  if (retained === null && affected.length === 0) {
    return;
  }
  appendProof(transaction, {
    type: "replace-selection-then-split-right-retains",
    affected: Object.freeze(affected),
    retained,
  });
};

type OwnershipException =
  | { token: string; type: "added" }
  | { token: string; type: "removed" }
  | { token: string; type: "relocated" };

const ownershipExceptions = (
  before: OwnershipSnapshot,
  after: OwnershipSnapshot,
  transaction: Transaction,
): OwnershipException[] => {
  const exceptions: OwnershipException[] = [];
  for (const [token, occurrence] of before.imported) {
    const next = after.imported.get(token);
    const mappedOwner = transaction.mapping.mapResult(occurrence.pos);
    const mappedInterior = transaction.mapping.mapResult(occurrence.pos + 1, -1);
    if (!next) {
      exceptions.push({ token, type: "removed" });
      continue;
    }
    if (mappedOwner.pos !== next.pos || mappedInterior.deletedAcross) {
      exceptions.push({ token, type: "relocated" });
    }
  }
  for (const token of after.imported.keys()) {
    if (!before.imported.has(token)) {
      exceptions.push({ token, type: "added" });
    }
  }
  return exceptions;
};

const proofExceptions = (
  transaction: Transaction,
  before: OwnershipSnapshot,
  after: OwnershipSnapshot,
): ReadonlySet<string> => {
  const permitted = new Set<string>();
  const originalSide = (direction: "forward" | "inverse") =>
    direction === "forward" ? before : after;
  const assertOriginalOccurrence = (
    direction: "forward" | "inverse",
    occurrence: ImportedOccurrence,
    message: string,
  ): void => {
    const actual = originalSide(direction).imported.get(occurrence.token);
    if (!actual || !importedOccurrenceMatches(actual, occurrence)) {
      panic(message);
    }
  };
  const retainedTransitionIsBounded = (
    direction: "forward" | "inverse",
    occurrence: ImportedOccurrence,
    sourceRange: { from: number; to: number } = {
      from: occurrence.pos,
      to: occurrence.pos + occurrence.nodeSize,
    },
  ): boolean => {
    if (direction === "inverse") {
      const restored = after.imported.get(occurrence.token);
      return restored !== undefined && importedOccurrenceMatches(restored, occurrence);
    }
    const source = before.imported.get(occurrence.token);
    const target = after.imported.get(occurrence.token);
    if (!source || !target || !importedOccurrenceMatches(source, occurrence)) {
      return false;
    }
    const mappedStart = transaction.mapping.map(sourceRange.from, -1);
    const mappedEnd = transaction.mapping.map(sourceRange.to, 1);
    return (
      target.pos >= Math.min(mappedStart, mappedEnd) &&
      target.pos <= Math.max(mappedStart, mappedEnd)
    );
  };
  const proofs = transaction.steps.filter(
    (step): step is ParagraphPropertyOwnershipTransitionStep =>
      step instanceof ParagraphPropertyOwnershipTransitionStep,
  );
  for (const proof of proofs) {
    const { direction, transition } = proof.read();
    switch (transition.type) {
      case "split-left-created-right-retains": {
        if (transition.retained === null) {
          break;
        }
        assertOriginalOccurrence(
          direction,
          transition.retained,
          "A paragraph split proof does not match its source owner",
        );
        if (!retainedTransitionIsBounded(direction, transition.retained)) {
          panic("A paragraph split proof does not match its retained source owner");
        }
        permitted.add(`relocated:${transition.retained.token}`);
        break;
      }
      case "join-left-paragraph-mark-retains":
      case "join-right-paragraph-mark-retains": {
        if (transition.retained !== null) {
          assertOriginalOccurrence(
            direction,
            transition.retained,
            "A paragraph join proof does not match its source owner",
          );
          if (
            !retainedTransitionIsBounded(direction, transition.retained, {
              from: transition.sourceFrom,
              to: transition.sourceTo,
            })
          ) {
            panic("A paragraph join proof does not match its retained source owner");
          }
          permitted.add(`relocated:${transition.retained.token}`);
        }
        if (transition.displaced !== null) {
          assertOriginalOccurrence(
            direction,
            transition.displaced,
            "A paragraph join proof does not match its displaced source owner",
          );
          const displacedBefore = before.imported.has(transition.displaced.token);
          const displacedAfter = after.imported.has(transition.displaced.token);
          if (direction === "forward" && (!displacedBefore || displacedAfter)) {
            panic("A paragraph join proof does not match its displaced source owner");
          }
          if (direction === "inverse" && (displacedBefore || !displacedAfter)) {
            panic("An inverse paragraph join proof did not restore its displaced source owner");
          }
          permitted.add(
            `${direction === "forward" ? "removed" : "added"}:${transition.displaced.token}`,
          );
        }
        break;
      }
      case "delete-selection":
        for (const occurrence of transition.affected) {
          assertOriginalOccurrence(
            direction,
            occurrence,
            "A paragraph deletion proof does not match its source owner",
          );
          const existsBefore = before.imported.has(occurrence.token);
          const existsAfter = after.imported.has(occurrence.token);
          if (direction === "forward" && !existsBefore) {
            panic("A paragraph deletion proof does not match its source owner");
          }
          if (direction === "inverse" && !existsAfter) {
            panic("An inverse paragraph deletion proof did not restore its source owner");
          }
          if (direction === "forward" && !existsAfter) {
            permitted.add(`removed:${occurrence.token}`);
          }
          if (direction === "inverse" && !existsBefore) {
            permitted.add(`added:${occurrence.token}`);
          }
          permitted.add(`relocated:${occurrence.token}`);
        }
        break;
      case "replace-selection-then-split-right-retains": {
        if (transition.retained !== null) {
          assertOriginalOccurrence(
            direction,
            transition.retained,
            "A replace-and-split proof does not match its source owner",
          );
          if (!retainedTransitionIsBounded(direction, transition.retained)) {
            panic("A replace-and-split proof lost its retained source owner");
          }
          permitted.add(`relocated:${transition.retained.token}`);
        }
        for (const occurrence of transition.affected) {
          assertOriginalOccurrence(
            direction,
            occurrence,
            "A replace-and-split proof does not match its source owner",
          );
          const existsBefore = before.imported.has(occurrence.token);
          const existsAfter = after.imported.has(occurrence.token);
          if (direction === "forward" && !existsBefore) {
            panic("A replace-and-split proof does not match its source owner");
          }
          if (direction === "inverse" && !existsAfter) {
            panic("An inverse replace-and-split proof did not restore its source owner");
          }
          if (occurrence.token === transition.retained?.token) {
            continue;
          }
          if (direction === "forward" && !existsAfter) {
            permitted.add(`removed:${occurrence.token}`);
          }
          if (direction === "inverse" && !existsBefore) {
            permitted.add(`added:${occurrence.token}`);
          }
          permitted.add(`relocated:${occurrence.token}`);
        }
        break;
      }
      default: {
        const exhaustive: never = transition;
        return exhaustive;
      }
    }
  }
  return permitted;
};

const isYjsTransaction = (transaction: Transaction): boolean =>
  transaction.getMeta(ySyncPluginKey) !== undefined;

const adjacentParagraphBefore = (
  snapshot: OwnershipSnapshot,
  target: ParagraphOccurrence,
): ParagraphOccurrence | undefined =>
  snapshot.paragraphs.find((paragraph) => paragraph.pos + paragraph.node.nodeSize === target.pos);

const remoteSplitIsProven = (
  token: string,
  transaction: Transaction,
  before: OwnershipSnapshot,
  after: OwnershipSnapshot,
): boolean => {
  const source = before.imported.get(token);
  const target = after.imported.get(token);
  if (!source || !target) {
    return false;
  }
  const targetParagraph = after.paragraphs.find((paragraph) => paragraph.pos === target.pos);
  if (!targetParagraph) {
    return false;
  }
  const left = adjacentParagraphBefore(after, targetParagraph);
  if (!left || left.state.type !== "editor-created") {
    return false;
  }
  if (!source.node.content.eq(left.node.content.append(target.node.content))) {
    return false;
  }
  return (
    transaction.mapping.map(source.pos, -1) === left.pos &&
    transaction.mapping.map(source.pos + source.nodeSize, 1) === target.pos + target.nodeSize
  );
};

const remoteJoinProofs = (
  transaction: Transaction,
  before: OwnershipSnapshot,
  after: OwnershipSnapshot,
): ReadonlySet<string> => {
  const permitted = new Set<string>();
  for (let index = 0; index + 1 < before.paragraphs.length; index += 1) {
    const left = before.paragraphs[index];
    const right = before.paragraphs[index + 1];
    if (!left || !right || left.pos + left.node.nodeSize !== right.pos) {
      continue;
    }
    const leftToken = left.state.type === "imported" ? left.state.token : null;
    const rightToken = right.state.type === "imported" ? right.state.token : null;
    if (leftToken === null && rightToken === null) {
      continue;
    }
    const mappedStart = transaction.mapping.map(left.pos, -1);
    const mappedEnd = transaction.mapping.map(right.pos + right.node.nodeSize, 1);
    const joinedCandidates = after.paragraphs.filter(
      (paragraph) =>
        paragraph.pos >= Math.min(mappedStart, mappedEnd) &&
        paragraph.pos <= Math.max(mappedStart, mappedEnd) &&
        paragraph.node.content.eq(left.node.content.append(right.node.content)),
    );
    const joined = joinedCandidates.length === 1 ? joinedCandidates.at(0) : undefined;
    if (!joined || joined.state.type !== "imported") {
      continue;
    }
    const retainedToken = joined.state.token;
    if (retainedToken !== leftToken && retainedToken !== rightToken) {
      continue;
    }
    permitted.add(`relocated:${retainedToken}`);
    const displacedToken = retainedToken === leftToken ? rightToken : leftToken;
    if (displacedToken !== null) {
      permitted.add(`removed:${displacedToken}`);
    }
  }
  return permitted;
};

const remoteTransitionProofs = (
  exceptions: readonly OwnershipException[],
  transaction: Transaction,
  before: OwnershipSnapshot,
  after: OwnershipSnapshot,
  knownOwners: ReadonlyMap<string, ImportedOccurrence>,
): ReadonlySet<string> => {
  const permitted = new Set(remoteJoinProofs(transaction, before, after));
  for (const exception of exceptions) {
    if (exception.type === "added") {
      const known = knownOwners.get(exception.token);
      const restored = after.imported.get(exception.token);
      if (
        known &&
        restored &&
        known.node.eq(restored.node) &&
        (transaction.mapping.map(known.pos, -1) === restored.pos ||
          transaction.mapping.map(known.pos, 1) === restored.pos)
      ) {
        permitted.add(`added:${exception.token}`);
      }
      continue;
    }
    if (exception.type === "removed") {
      const source = before.imported.get(exception.token);
      if (source && transaction.mapping.mapResult(source.pos + 1, -1).deletedAcross) {
        permitted.add(`removed:${exception.token}`);
      }
      continue;
    }
    if (
      exception.type === "relocated" &&
      remoteSplitIsProven(exception.token, transaction, before, after)
    ) {
      permitted.add(`relocated:${exception.token}`);
    }
  }
  return permitted;
};

const advanceKnownOwners = (
  knownOwners: ReadonlyMap<string, ImportedOccurrence>,
  after: OwnershipSnapshot,
  transaction: Transaction,
): ReadonlyMap<string, ImportedOccurrence> => {
  const next = new Map<string, ImportedOccurrence>();
  for (const [token, known] of knownOwners) {
    const current = after.imported.get(token);
    next.set(
      token,
      current ??
        Object.freeze({
          ...known,
          pos: transaction.mapping.map(known.pos, -1),
        }),
    );
  }
  for (const [token, current] of after.imported) {
    if (!next.has(token)) {
      next.set(token, current);
    }
  }
  return next;
};

/** Nominal transaction context created by a complete PM/Yjs ownership census. */
export class ParagraphPropertyDocumentContext {
  readonly #contract: ParagraphPropertySourceContract | null;
  readonly #knownOwners: ReadonlyMap<string, ImportedOccurrence>;
  readonly #snapshot: OwnershipSnapshot;
  readonly #story: ParagraphPropertySourceStory;

  private constructor(
    contract: ParagraphPropertySourceContract | null,
    snapshot: OwnershipSnapshot,
    knownOwners: ReadonlyMap<string, ImportedOccurrence>,
    story: ParagraphPropertySourceStory,
  ) {
    this.#contract = contract;
    this.#snapshot = snapshot;
    this.#knownOwners = knownOwners;
    this.#story = story;
    Object.freeze(this);
  }

  static fromDocument(
    doc: PMNode,
    story: ParagraphPropertySourceStory,
  ): ParagraphPropertyDocumentContext {
    const contract = readContract(doc);
    const snapshot = ownershipSnapshot(doc, contract, story);
    return new ParagraphPropertyDocumentContext(
      contract,
      snapshot,
      new Map(snapshot.imported),
      story,
    );
  }

  advance(transaction: Transaction): ParagraphPropertyDocumentContext {
    const contract = readContract(transaction.doc);
    if (contract?.serialized !== this.#contract?.serialized) {
      throw new ParagraphPropertySourceValidationError({
        code: "contract_mismatch",
        message: "A transaction changed its paragraph-property source contract.",
      });
    }
    const after = ownershipSnapshot(transaction.doc, contract, this.#story);
    const exceptions = ownershipExceptions(this.#snapshot, after, transaction);
    const proofs = proofExceptions(transaction, this.#snapshot, after);
    const yjs = isYjsTransaction(transaction);
    const remoteProofs = yjs
      ? remoteTransitionProofs(exceptions, transaction, this.#snapshot, after, this.#knownOwners)
      : new Set<string>();
    for (const exception of exceptions) {
      if (proofs.has(`${exception.type}:${exception.token}`)) {
        continue;
      }
      if (remoteProofs.has(`${exception.type}:${exception.token}`)) {
        continue;
      }
      throw new ParagraphPropertySourceValidationError({
        code: "ownership_transition_mismatch",
        message: yjs
          ? "A remote paragraph-property ownership transition is ambiguous."
          : "A local paragraph-property ownership transition lacks an issuer-minted proof.",
        token: exception.token,
      });
    }
    return new ParagraphPropertyDocumentContext(
      contract,
      after,
      advanceKnownOwners(this.#knownOwners, after, transaction),
      this.#story,
    );
  }
}
