/**
 * One run identity for the pieces of a `w:r` a tracked change cut.
 *
 * A revision inside a run splits it: `<w:r a>Sel</w:r><w:del><w:r a>ler</w:r></w:del>`,
 * `<w:r a>Sel</w:r><w:ins>…</w:ins><w:r a>ler</w:r>`, or a `w:rPrChange` on
 * one piece of the run only. Read back, each
 * `w:r` is its own authored run, so resolving the revision would leave two
 * runs where the source had one. The package cannot say whether the two were
 * one run, but runs that carry the same attributes and that only revised
 * content holds apart are what such a cut writes, so the later piece takes the
 * earlier one's identity. The pieces stay separate runs while the revision
 * holds them apart, and become one run once it is resolved and they state the
 * same properties again (a run joins its owner only then, see `fromProseDoc`).
 *
 * Runs the package writes side by side under the same revision state keep
 * their own identities.
 */

import { Fragment, Mark, type Node as PMNode } from "prosemirror-model";

import { expectRunIdentityMarkAttrs } from "./attrs";
import { INLINE_CONTENT_CONTROL_NODE_NAME } from "./extensions/nodes/SdtExtension";
import { RUN_IDENTITY_MARK_NAME } from "./runIdentity";

const REVISION_MARK_NAMES: ReadonlySet<string> = new Set([
  "insertion",
  "deletion",
  "runPropertyChange",
]);

const identityOf = (node: PMNode): Mark | undefined =>
  node.marks.find((mark) => mark.type.name === RUN_IDENTITY_MARK_NAME);

const revisionMarksOf = (node: PMNode): readonly Mark[] =>
  node.marks.filter((mark) => REVISION_MARK_NAMES.has(mark.type.name));

/** The identity's attrs without the id: what the run carries. */
const payloadKey = (identity: Mark): string =>
  JSON.stringify({ ...expectRunIdentityMarkAttrs(identity), id: 0 });

type ContinuesRunOptions = {
  left: PMNode;
  right: PMNode;
  /** Whether revised content lies between the two. */
  bridged: boolean;
};

/** Whether `right` continues `left`'s run across a revision boundary. */
const continuesRun = ({ left, right, bridged }: ContinuesRunOptions): boolean => {
  const leftIdentity = identityOf(left);
  const rightIdentity = identityOf(right);
  if (!leftIdentity || !rightIdentity || leftIdentity.eq(rightIdentity)) {
    return false;
  }
  return (
    (bridged || !Mark.sameSet(revisionMarksOf(left), revisionMarksOf(right))) &&
    payloadKey(leftIdentity) === payloadKey(rightIdentity)
  );
};

/**
 * `nodes`, one inline container's content, with each run a revision cut given
 * the identity of the run it was cut from. A control is a container of its
 * own. `nodes` itself comes back when nothing changes.
 */
export const continueRunIdentitiesAcrossRevisions = (
  nodes: readonly PMNode[],
): readonly PMNode[] => {
  const renamed = new Map<number, Mark>();
  /** The last run with an identity, while only revised content follows it. */
  let anchor: PMNode | undefined;
  let bridged = false;
  let changed = false;
  const result = nodes.map((node) => {
    if (node.type.name === INLINE_CONTENT_CONTROL_NODE_NAME) {
      anchor = undefined;
      const children: PMNode[] = [];
      node.forEach((child) => {
        children.push(child);
      });
      const continued = continueRunIdentitiesAcrossRevisions(children);
      if (continued === children) {
        return node;
      }
      changed = true;
      return node.copy(Fragment.fromArray([...continued]));
    }
    const identity = identityOf(node);
    if (!identity) {
      if (anchor && revisionMarksOf(node).length > 0) {
        bridged = true;
      } else {
        anchor = undefined;
      }
      return node;
    }
    const id = expectRunIdentityMarkAttrs(identity).id;
    let replacement = renamed.get(id);
    if (!replacement && anchor && continuesRun({ left: anchor, right: node, bridged })) {
      replacement = identityOf(anchor);
      if (replacement) {
        renamed.set(id, replacement);
      }
    }
    const next = replacement ? node.mark(replacement.addToSet(node.marks)) : node;
    changed ||= next !== node;
    anchor = next;
    bridged = false;
    return next;
  });
  return changed ? result : nodes;
};
