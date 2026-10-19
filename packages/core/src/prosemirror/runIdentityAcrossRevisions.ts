/**
 * One source identity for matching runs and hyperlinks a revision cut apart.
 * Serialized fragments mint separate identities on reopen. Matching fragments
 * across a revision boundary or revised-only bridge take the earlier identity,
 * so resolution can rejoin them. Every non-identity attribute must match.
 * Adjacent carriers under the same revision state keep their source identities.
 */

import { Fragment, Mark, type Node as PMNode } from "prosemirror-model";

import { expectRunIdentityMarkAttrs, expectHyperlinkMarkAttrs } from "./attrs";
import { INLINE_CONTENT_CONTROL_NODE_NAME } from "./extensions/nodes/SdtExtension";
import { RUN_IDENTITY_MARK_NAME } from "./runIdentity";

const REVISION_MARK_NAMES: ReadonlySet<string> = new Set([
  "insertion",
  "deletion",
  "runPropertyChange",
]);

/** Ordered source of truth for source-carrier identities and continuation. */
export const SOURCE_CARRIER_MARK_NAMES = [RUN_IDENTITY_MARK_NAME, "hyperlink"] as const;
type SourceCarrierMarkName = (typeof SOURCE_CARRIER_MARK_NAMES)[number];
type CarrierIdentity = { id: number; payloadKey: string } | undefined;
type CarrierOwner<Name extends SourceCarrierMarkName = SourceCarrierMarkName> = {
  markName: Name;
  read: (mark: Mark) => CarrierIdentity;
};
type CarrierOwners = { readonly [Name in SourceCarrierMarkName]: CarrierOwner<Name> };

/** Every declared carrier has one owner; iteration follows the canonical tuple. */
const CARRIER_OWNERS = {
  [SOURCE_CARRIER_MARK_NAMES[0]]: {
    markName: SOURCE_CARRIER_MARK_NAMES[0],
    read: (mark: Mark) => {
      const attrs = expectRunIdentityMarkAttrs(mark);
      return { id: attrs.id, payloadKey: JSON.stringify({ ...attrs, id: 0 }) };
    },
  },
  [SOURCE_CARRIER_MARK_NAMES[1]]: {
    markName: SOURCE_CARRIER_MARK_NAMES[1],
    read: (mark: Mark) => {
      const attrs = expectHyperlinkMarkAttrs(mark);
      if (typeof attrs._docxHyperlinkIndex !== "number") return undefined;
      return {
        id: attrs._docxHyperlinkIndex,
        payloadKey: JSON.stringify({ ...attrs, _docxHyperlinkIndex: 0 }),
      };
    },
  },
} satisfies CarrierOwners;
type CarrierOwnerEntry = CarrierOwners[SourceCarrierMarkName];

const identityOf = (node: PMNode, owner: CarrierOwnerEntry) => {
  const mark = node.marks.find((candidate) => candidate.type.name === owner.markName);
  if (!mark) return undefined;
  const identity = owner.read(mark);
  return identity === undefined ? undefined : { mark, ...identity };
};

const revisionMarksOf = (node: PMNode): readonly Mark[] =>
  node.marks.filter((mark) => REVISION_MARK_NAMES.has(mark.type.name));

type ContinuesCarrierOptions = {
  left: PMNode;
  right: PMNode;
  /** Whether revised content lies between the two. */
  bridged: boolean;
  owner: CarrierOwnerEntry;
};

/** Whether `right` continues `left`'s run across a revision boundary. */
const continuesCarrier = ({ left, right, bridged, owner }: ContinuesCarrierOptions): boolean => {
  const leftIdentity = identityOf(left, owner);
  const rightIdentity = identityOf(right, owner);
  if (!leftIdentity || !rightIdentity || leftIdentity.mark.eq(rightIdentity.mark)) {
    return false;
  }
  return (
    (bridged || !Mark.sameSet(revisionMarksOf(left), revisionMarksOf(right))) &&
    leftIdentity.payloadKey === rightIdentity.payloadKey
  );
};

/**
 * `nodes`, one inline container's content, with each revision fragment given
 * the identity of its matching source carrier. A control is a container of its
 * own. `nodes` itself comes back when nothing changes.
 */
const continueIdentities = (
  nodes: readonly PMNode[],
  owner: CarrierOwnerEntry,
): readonly PMNode[] => {
  const renamed = new Map<number, Mark>();
  /** The last carrier with an identity, while only revised content follows it. */
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
      const continued = continueIdentities(children, owner);
      if (continued === children) {
        return node;
      }
      changed = true;
      return node.copy(Fragment.fromArray([...continued]));
    }
    const identity = identityOf(node, owner);
    if (!identity) {
      if (anchor && revisionMarksOf(node).length > 0) {
        bridged = true;
      } else {
        anchor = undefined;
      }
      return node;
    }
    const id = identity.id;
    let replacement = renamed.get(id);
    if (!replacement && anchor && continuesCarrier({ left: anchor, right: node, bridged, owner })) {
      replacement = identityOf(anchor, owner)?.mark;
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

/** Every source-carrier family uses the same revision-boundary continuation policy. */
export const continueSourceCarrierIdentitiesAcrossRevisions = (
  nodes: readonly PMNode[],
): readonly PMNode[] => {
  let continued = nodes;
  for (const markName of SOURCE_CARRIER_MARK_NAMES) {
    continued = continueIdentities(continued, CARRIER_OWNERS[markName]);
  }
  return continued;
};
