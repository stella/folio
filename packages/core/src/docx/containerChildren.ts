/**
 * The shared child dispatcher: one way to walk a container's children.
 *
 * A hand-rolled `switch` over child local names decides two things at once —
 * which children are modelled, and what happens to the rest — and it only ever
 * writes the first down. The second is the `default` branch, and a `default`
 * that does nothing is how folio loses markup silently.
 *
 * Here the two are separated. The handler map is **total** over the children
 * the container's content model declares (`containerChildren.gen.ts`, derived
 * from the committed schema graph), so a container cannot gain a declared
 * child without somebody choosing `CAPTURE`, `ownedElsewhere`, or a handler:
 * the compiler refuses the map otherwise. And anything *undeclared* — a
 * foreign namespace, an `mc:` construct, an element a later OOXML revision
 * adds — goes to the ordered verbatim sink by default, because a name the
 * schema does not know is exactly the case a hand-written `default` gets
 * wrong.
 *
 * The sink records position as an ordinal rather than a pointer, so
 * `serializeWithPreservedChildren` can put the markup back between the same
 * neighbours. See `preservedMarkup.ts` for why. What the ordinal counts is the
 * caller's to decide: a count of modelled siblings where the container models
 * one kind of child, and the schema's own sequence position where its content
 * model is a fixed sequence — see {@link sequencePositions}.
 */

import type { PreservedChild, PreservedMarkup } from "@stll/docx-core/model";

import {
  CONTAINER_CHILDREN,
  type DeclaredChild,
  type DispatchedContainer,
  type SequenceContainer,
} from "./containerChildren.gen";
import { TRANSITIONAL_NAMESPACE_BY_STRICT_URI } from "./strictValueEncodings.gen";
import { captureVerbatimXml } from "./verbatimCapture";
import {
  getChildElements,
  getLocalName,
  getNamespaceUri,
  WORDPROCESSINGML_NAMESPACE_URIS,
  type XmlElement,
} from "./xmlParser";

/** Keep this child's markup verbatim, in place. */
export const CAPTURE = "capture";

/**
 * What a handler answers with: nothing when it took a typed value, and
 * {@link CAPTURE} when it took none, so the child's bytes are kept instead.
 *
 * `<w:tblLayout/>` states no layout and `<w:jc w:val="end"/>` states a value a
 * reader's enumeration does not admit. Both used to fall off the end of the
 * walk, and neither can be decided by name: a map keyed by name cannot list
 * the values a reader will refuse, only the names it has never heard of.
 */
export const keptUnless = (taken: boolean): typeof CAPTURE | undefined =>
  taken ? undefined : CAPTURE;

const OWNED_ELSEWHERE = "owned-elsewhere";

/**
 * Who reads a child this walk skips: `<module>#<export>`.
 *
 * The module is a file under `packages/core/src/docx/` without its extension,
 * and the export is the function that reads the child off the container
 * element. Both halves are resolved by `scripts/container-ownership.test.ts`,
 * which imports every one of them.
 */
export type ChildOwnerRef = `${string}#${string}`;

/**
 * Another reader owns this child, and re-emits it.
 *
 * `w:rPr` under a run, read by `parseRunProperties`; `w:commentReference`
 * under a run, lifted out by the paragraph parser. Capturing one of these as
 * well would write it twice, so the disposition has to be stated rather than
 * left to a bare `break`.
 *
 * A capture and a handler are verified by running them: the survival law
 * exercises the pair and reports what came back. This one is a claim about a
 * *different module*, and nothing ran it. `w:tr/w:tblPrEx` carried it while
 * the container contract recorded the same pair as never parsed, and the two
 * statements sat side by side without contradicting each other. So the claim
 * names its owner, `ownedElsewhere` registers it where a check can read it,
 * and `scripts/lib/container-survival/ownership.ts` refuses a contract that
 * records an owned pair as never parsed or lost with its container.
 *
 * The container and the child are on the value rather than inferred from the
 * key, and {@link ChildHandlers} binds them to it: an entry that names a child
 * other than the one it is filed under does not compile.
 */
export type OwnedElsewhere<
  Container extends DispatchedContainer = DispatchedContainer,
  Child extends string = string,
> = {
  readonly disposition: typeof OWNED_ELSEWHERE;
  readonly container: Container;
  readonly child: Child;
  readonly reader: ChildOwnerRef;
};

/** One handler map's claim that another reader owns one of its children. */
export type ChildOwnerClaim = {
  readonly container: DispatchedContainer;
  readonly child: string;
  readonly reader: ChildOwnerRef;
};

const CLAIMS = new Map<string, ChildOwnerClaim>();

/**
 * Every owner a handler map has claimed so far.
 *
 * A claim is registered when `ownedElsewhere` is called, so a module that is
 * never imported contributes none. The checks that read this therefore load
 * the claiming modules first, by scanning the sources for the call rather than
 * by keeping a list beside them.
 */
export const ownedElsewhereClaims = (): readonly ChildOwnerClaim[] => [...CLAIMS.values()];

/** State that another reader takes this child, and who. */
export const ownedElsewhere = <
  Container extends DispatchedContainer,
  Child extends DeclaredChild<Container>,
>(
  claim: ChildOwnerClaim & { container: Container; child: Child },
): OwnedElsewhere<Container, Child> => {
  CLAIMS.set(`${claim.container}/${claim.child}#${claim.reader}`, claim);
  return { disposition: OWNED_ELSEWHERE, ...claim };
};

/**
 * This child goes with the wrapper folio does not keep.
 *
 * `w:smartTagPr` is the case: folio unwraps `w:smartTag` and splices its
 * content into the paragraph, so the properties describing the wrapper have
 * nothing left to describe and capturing them would put a `w:smartTagPr`
 * where the schema does not admit one. The drop is stated here and recorded
 * in `specifications/container-contract/contract.json`, which is the whole
 * difference between this and a `default` that says nothing.
 */
export const DROPPED_WITH_ITS_WRAPPER = "dropped-with-its-wrapper";

/**
 * What a container does with one declared child.
 *
 * A handler that reads the child may hand it back by returning {@link CAPTURE}
 * — "I looked and took nothing from this, so keep the bytes". A property set
 * needs it: `<w:cols/>` states no column count, `<w:jc w:val="end"/>` states a
 * value the reader's enumeration does not admit, and a handler that turns
 * neither into a typed value leaves the element with nowhere to go. Deciding
 * it by the outcome rather than by the name is what makes the decision total:
 * the map cannot list the values a reader will refuse.
 */
export type ChildDisposition<
  Container extends DispatchedContainer = DispatchedContainer,
  Child extends string = string,
> =
  | ((child: XmlElement) => typeof CAPTURE | void)
  | typeof CAPTURE
  | typeof DROPPED_WITH_ITS_WRAPPER
  | OwnedElsewhere<Container, Child>;

/**
 * A total decision per declared child.
 *
 * Total, never `Partial`: a partial map lets a new declared child land without
 * a decision, which is the drop this module exists to prevent. The value's
 * type is keyed by the child it is filed under, so an {@link ownedElsewhere}
 * entry cannot name a different child than the one it decides.
 */
export type ChildHandlers<Container extends DispatchedContainer> = Readonly<{
  [Child in DeclaredChild<Container>]: ChildDisposition<Container, Child>;
}>;

type DispatchChildrenOptions<Container extends DispatchedContainer> = {
  /** The container element, as the source wrote it. */
  element: XmlElement;
  /** Which container's declared-child set the handler map must be total over. */
  container: Container;
  handlers: ChildHandlers<Container>;
  /**
   * Where a capture belongs among the caller's children, called once per
   * captured child.
   *
   * A container that models one kind of child answers with how many it holds
   * right now, so the capture lands after the siblings already read. A
   * container whose content model is a fixed sequence answers with the child's
   * own ordinal in that sequence, because there the position is decided by the
   * name rather than by what happened to be read first; see
   * {@link sequencePositions}.
   */
  capturePosition: (child: XmlElement) => number;
  /**
   * Dispositions for names the container's content model does not declare.
   *
   * The schema declares a container's children in one namespace, so anything
   * else is undeclared by construction — which is why the sink is the default
   * and why this map is not part of the totality check: there is no finite set
   * to be total over. `mc:AlternateContent` is the case that needs it. It is
   * markup compatibility, legal wherever its fallback is, and folio models it
   * by selecting a branch and reading that; capturing the wrapper whole would
   * keep the bytes and lose every paragraph inside it to the editor.
   *
   * Each entry is a claim that folio reads this name in this container.
   */
  undeclared?: Readonly<Record<string, ChildDisposition>>;
  /**
   * Dispositions for whole namespaces the declared set does not cover, by
   * namespace URI, consulted after {@link undeclared} and before the sink.
   *
   * OOXML maths is the case. `m:EG_OMathMathElements` is admitted wherever
   * `m:oMath` is, so `<w:ins><m:f/></w:ins>` is a tracked insertion of a
   * fraction with no `m:oMath` around it; there are two dozen such names and
   * folio treats every one of them the same way, as markup it carries rather
   * than a structure it models. Naming the namespace states that once. It is
   * not a `default` in disguise: a child from any *other* namespace still goes
   * to the sink, which is the branch a hand-written `default` gets wrong.
   */
  undeclaredNamespaces?: Readonly<Record<string, ChildDisposition>>;
};

/**
 * The Transitional spelling of a namespace, so one disposition covers both.
 *
 * ISO Strict writes the same content model under its own URIs: maths is
 * `purl.oclc.org/ooxml/officeDocument/math` rather than
 * `schemas.openxmlformats.org/officeDocument/2006/math`. A namespace-keyed
 * disposition written in Transitional would otherwise send a Strict
 * document's maths to the sink, which models it as bytes instead of an
 * equation. The pairs come from the generated table rather than a second list
 * here, so a namespace cannot be paired in one place and forgotten in the
 * other.
 */
export const transitionalNamespaceOf = (namespace: string): string =>
  TRANSITIONAL_NAMESPACE_BY_STRICT_URI.get(namespace) ?? namespace;

/**
 * Walk a container's children, handing each to its handler and the rest to the
 * sink.
 *
 * @returns the container's unmodelled markup, or `undefined` when it held
 *   none — an empty record is never written, so a fully modelled container
 *   stays byte-identical in the model.
 */
export const dispatchChildren = <Container extends DispatchedContainer>({
  element,
  handlers,
  capturePosition,
  undeclared,
  undeclaredNamespaces,
}: DispatchChildrenOptions<Container>): PreservedMarkup | undefined => {
  const children: PreservedChild[] = [];
  const capture = (child: XmlElement): void => {
    children.push({ index: capturePosition(child), xml: captureVerbatimXml(child) });
  };

  const declared = new Map<string, ChildDisposition>(Object.entries(handlers));
  const byName = new Map<string, ChildDisposition>(Object.entries(undeclared ?? {}));
  const byNamespace = new Map<string, ChildDisposition>(Object.entries(undeclaredNamespaces ?? {}));
  for (const child of getChildElements(element)) {
    // The generated set names one namespace's children, so a child is
    // declared only when it is in that namespace. Matching on the local name
    // alone hands `m:r` to the `w:r` handler, and a math run parsed as a text
    // run comes out empty and is pruned: the collision is a silent loss with
    // a handler in front of it, which is harder to see than no handler at all.
    // An `undeclared` name is matched across namespaces, because the entries
    // that need it — `mc:AlternateContent` above all — are named by a prefix
    // the container's own namespace never binds.
    const localName = getLocalName(child.name);
    const namespace = getNamespaceUri(child);
    // An element whose prefix nothing in scope binds has no namespace to be
    // told apart by, and folio's other readers match it by prefix alone. The
    // collision this guards against needs both namespaces bound to exist, so
    // an unresolved child is read as the container's own, exactly as before.
    const isDeclaredNamespace =
      namespace === undefined || WORDPROCESSINGML_NAMESPACE_URIS.has(namespace);
    const disposition = isDeclaredNamespace
      ? (declared.get(localName) ?? byName.get(localName))
      : (byName.get(localName) ?? byNamespace.get(transitionalNamespaceOf(namespace)));
    if (disposition === undefined || disposition === CAPTURE) {
      capture(child);
      continue;
    }
    // `DROPPED_WITH_ITS_WRAPPER`, or a child another reader owns: either way
    // this walk writes nothing down and the decision is recorded elsewhere.
    if (typeof disposition !== "function") {
      continue;
    }
    if (disposition(child) === CAPTURE) {
      capture(child);
    }
  }

  return children.length === 0 ? undefined : { children };
};

/**
 * Put a container's modelled children back with its preserved ones between
 * them, in source position.
 *
 * Where the content model demands schema order the caller has already put its
 * modelled children in that order; the sink's `index` then places each capture
 * between the same two siblings it sat between when it was read.
 *
 * @param wrap turns one capture's markup into whatever the caller's list holds
 */
export const withPreservedChildren = <Item>(
  modelled: readonly Item[],
  preserved: PreservedMarkup | undefined,
  wrap: (xml: string) => Item,
): Item[] => {
  const captures = preserved?.children;
  if (captures === undefined || captures.length === 0) {
    return [...modelled];
  }

  const byIndex = new Map<number, Item[]>();
  for (const { index, xml } of captures) {
    // A capture recorded past the modelled count — the model lost the sibling
    // it followed — lands at the end rather than being dropped.
    const slot = Math.min(Math.max(index, 0), modelled.length);
    const bucket = byIndex.get(slot);
    if (bucket) {
      bucket.push(wrap(xml));
      continue;
    }
    byIndex.set(slot, [wrap(xml)]);
  }

  const items: Item[] = [];
  for (const [index, item] of modelled.entries()) {
    items.push(...(byIndex.get(index) ?? []), item);
  }
  items.push(...(byIndex.get(modelled.length) ?? []));
  return items;
};

/** {@link withPreservedChildren} for a caller that already holds markup. */
export const serializeWithPreservedChildren = (
  modelled: readonly string[],
  preserved: PreservedMarkup | undefined,
): string => withPreservedChildren(modelled, preserved, (xml) => xml).join("");

/**
 * Where each child sits in a container whose content model is one flat sequence.
 *
 * A property set is that container: `CT_TblPr` and `CT_SectPr` are sequences
 * of optional singletons, so a child's position is a property of its name and
 * the generated list is the order. That is what the sink records for them,
 * rather than a count of modelled siblings: the count is a mirror of whichever
 * properties folio models today, and it moves under the capture the moment one
 * more of them is modelled.
 *
 * An undeclared child — a foreign namespace, an `mc:` construct, a name a
 * later revision of the format adds — has no place in the sequence, so it
 * takes the place of the last declared child before it and comes back beside
 * the same neighbour.
 */
export const sequencePositions = <Container extends SequenceContainer>(
  container: Container,
  element: XmlElement,
): ((child: XmlElement) => number) => {
  const declared: readonly string[] = CONTAINER_CHILDREN[container];
  const positions = new Map<XmlElement, number>();
  let previous = 0;
  for (const child of getChildElements(element)) {
    const at = declared.indexOf(getLocalName(child.name));
    if (at !== -1) {
      previous = at;
    }
    positions.set(child, at === -1 ? previous : at);
  }
  return (child) => positions.get(child) ?? 0;
};

/**
 * A sequence container's children, modelled and captured, in schema order.
 *
 * The modelled half is keyed by element name rather than pre-ordered by the
 * caller, so the order is read from the generated sequence instead of being
 * restated as the order of a list of `if` statements — the restatement is what
 * drifted from the schema before, and a consumer refuses a property set whose
 * children are out of order. The captured half carries the ordinal the parser
 * read it at (see {@link sequencePositions}), so the two merge by the same
 * key and a capture lands where the source put it.
 *
 * A sequence declares every child at most once, so ties are only possible
 * between a declared child and an undeclared one sharing its slot; the sort is
 * stable and the modelled child leads.
 */
export const serializeSequenceChildren = <Container extends SequenceContainer>({
  container,
  modelled,
  preserved,
}: {
  container: Container;
  modelled: ReadonlyArray<readonly [name: DeclaredChild<Container>, xml: string]>;
  preserved: PreservedMarkup | undefined;
}): string[] => {
  const declared: readonly string[] = CONTAINER_CHILDREN[container];
  const placed: Array<{ at: number; rank: number; xml: string }> = [];
  for (const [name, xml] of modelled) {
    if (xml.length > 0) {
      placed.push({ at: declared.indexOf(name), rank: 0, xml });
    }
  }
  for (const { index, xml } of preserved?.children ?? []) {
    placed.push({ at: index, rank: 1, xml });
  }
  return placed
    .sort((left, right) => left.at - right.at || left.rank - right.rank)
    .map(({ xml }) => xml);
};
