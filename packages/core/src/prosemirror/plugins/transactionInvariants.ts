/**
 * Transaction invariants: a development and test aid that checks the editor
 * document after every transaction, so a command that leaves a document the
 * save path will refuse fails where it happens rather than at save.
 *
 * Off by default. {@link ExtensionManager} installs it when
 * `globalThis.__folioTransactionInvariants` is `{ enabled: true }` at the time
 * the runtime is initialized; the core test preload sets that flag, which turns
 * every editor test built on the extension manager into a validity check. A
 * host may set it in a development build. Production code never reads it
 * otherwise.
 *
 * The checks are the ones the document-model validator makes that the editor
 * state can answer on its own:
 * - the ProseMirror document is structurally valid (node and mark attributes,
 *   bookmark pairing), incrementally cached per node;
 * - every paragraph numbering reference names an instance the document's
 *   numbering defines.
 */

import type { Node as PMNode } from "prosemirror-model";
import { Plugin, PluginKey } from "prosemirror-state";
import type { EditorState, Transaction } from "prosemirror-state";

import { paragraphNumberingReferenceId } from "../../docx/numberingReference";
import { validateProseMirrorDocument } from "../validation";
import { getDocumentNumbering, hasDocumentNumbering } from "./documentNumbering";

export type TransactionInvariantFlags = {
  /** Install the invariant plugin in every extension-manager runtime. */
  enabled?: boolean;
};

declare global {
  var __folioTransactionInvariants: TransactionInvariantFlags | undefined;
}

/** Whether new extension-manager runtimes install the invariant plugin. */
export const areTransactionInvariantsEnabled = (): boolean =>
  globalThis.__folioTransactionInvariants?.enabled === true;

export type TransactionInvariantIssue = {
  path: string;
  message: string;
  /** Position of the node the issue sits on, when the check knows it. */
  pos?: number;
};

const MAX_REPORTED_STEPS_LENGTH = 4000;

const describeSteps = (transactions: readonly Transaction[]): string => {
  const steps = transactions.flatMap((transaction) =>
    transaction.steps.map((step) => step.toJSON() as unknown),
  );
  const json = JSON.stringify(steps);
  return json.length > MAX_REPORTED_STEPS_LENGTH
    ? `${json.slice(0, MAX_REPORTED_STEPS_LENGTH)}… (${json.length} characters)`
    : json;
};

export class TransactionInvariantError extends Error {
  readonly issues: readonly TransactionInvariantIssue[];
  readonly steps: readonly unknown[];

  constructor(issues: readonly TransactionInvariantIssue[], transactions: readonly Transaction[]) {
    super(
      `A transaction left the editor document invalid:\n${issues
        .map((issue) => `- ${issue.path}: ${issue.message}`)
        .join("\n")}\nSteps: ${describeSteps(transactions)}`,
    );
    this.name = "TransactionInvariantError";
    this.issues = issues;
    this.steps = transactions.flatMap((transaction) =>
      transaction.steps.map((step) => step.toJSON() as unknown),
    );
  }
}

const numberingReferenceIssues = (state: EditorState): TransactionInvariantIssue[] => {
  if (!hasDocumentNumbering(state)) {
    return [];
  }
  const numbering = getDocumentNumbering(state);
  const issues: TransactionInvariantIssue[] = [];
  state.doc.descendants((node: PMNode, pos: number) => {
    if (node.type.name !== "paragraph") {
      return true;
    }
    const numId = paragraphNumberingReferenceId(
      node.attrs["numPr"] as Parameters<typeof paragraphNumberingReferenceId>[0],
    );
    if (numId !== undefined && !numbering?.hasNumbering(numId)) {
      issues.push({
        path: `paragraph at ${pos}.numPr.numId`,
        message: `Numbering definition ${numId} is missing.`,
        pos,
      });
    }
    return false;
  });
  return issues;
};

/** Every invariant the state breaks; empty when it is sound. */
export const checkEditorStateInvariants = (state: EditorState): TransactionInvariantIssue[] => [
  ...validateProseMirrorDocument(state.doc).issues,
  ...numberingReferenceIssues(state),
];

/**
 * The issues `after` has that `before` did not. A document can arrive broken
 * (a malformed attribute the editor tolerates); only what a transaction adds
 * is the transaction's fault. An issue that names its node persists only on
 * that node, followed through the transactions' mappings, so moving a broken
 * reference from one paragraph to another still counts as new. Issues without
 * a position are matched by message, since their paths move with the edit.
 */
const introducedIssues = (
  before: readonly TransactionInvariantIssue[],
  after: readonly TransactionInvariantIssue[],
  transactions: readonly Transaction[],
): TransactionInvariantIssue[] => {
  const mapPosition = (pos: number): number => {
    let mapped = pos;
    for (const transaction of transactions) {
      mapped = transaction.mapping.map(mapped, 1);
    }
    return mapped;
  };
  const persisting = new Set<string>();
  const remaining = new Map<string, number>();
  for (const issue of before) {
    if (issue.pos === undefined) {
      remaining.set(issue.message, (remaining.get(issue.message) ?? 0) + 1);
    } else {
      persisting.add(`${mapPosition(issue.pos)}:${issue.message}`);
    }
  }
  return after.filter((issue) => {
    if (issue.pos !== undefined) {
      return !persisting.has(`${issue.pos}:${issue.message}`);
    }
    const count = remaining.get(issue.message) ?? 0;
    if (count === 0) {
      return true;
    }
    remaining.set(issue.message, count - 1);
    return false;
  });
};

const transactionInvariantsKey = new PluginKey("transactionInvariants");

/**
 * Throw a {@link TransactionInvariantError} naming the offending steps when a
 * document-changing transaction leaves the state with an issue it did not
 * have before.
 */
export const createTransactionInvariantPlugin = (): Plugin =>
  new Plugin({
    key: transactionInvariantsKey,
    appendTransaction(transactions, oldState, newState) {
      if (!transactions.some((transaction) => transaction.docChanged)) {
        return null;
      }
      const issues = checkEditorStateInvariants(newState);
      if (issues.length === 0) {
        return null;
      }
      const introduced = introducedIssues(
        checkEditorStateInvariants(oldState),
        issues,
        transactions,
      );
      if (introduced.length > 0) {
        throw new TransactionInvariantError(introduced, transactions);
      }
      return null;
    },
  });
