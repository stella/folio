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
 * child without somebody choosing `CAPTURE`, `OWNED_ELSEWHERE`, or a handler:
 * the compiler refuses the map otherwise. And anything *undeclared* — a
 * foreign namespace, an `mc:` construct, an element a later OOXML revision
 * adds — goes to the ordered verbatim sink by default, because a name the
 * schema does not know is exactly the case a hand-written `default` gets
 * wrong.
 *
 * The sink records position as a count of modelled siblings rather than a
 * pointer, so `serializeWithPreservedChildren` can put the markup back between
 * the same neighbours. See `preservedMarkup.ts` for why.
 */

import type { PreservedChild, PreservedMarkup } from "@stll/docx-core/model";

import type { DeclaredChild, DispatchedContainer } from "./containerChildren.gen";
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
 * Another reader owns this child, and re-emits it.
 *
 * `w:rPr` under a run, read by `parseRunProperties`; `w:commentReference`
 * under a run, lifted out by the paragraph parser. Capturing one of these as
 * well would write it twice, so the disposition has to be stated rather than
 * left to a bare `break`.
 */
export const OWNED_ELSEWHERE = "owned-elsewhere";

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

/** What a container does with one declared child. */
export type ChildDisposition =
  | ((child: XmlElement) => void)
  | typeof CAPTURE
  | typeof DROPPED_WITH_ITS_WRAPPER
  | typeof OWNED_ELSEWHERE;

/**
 * A total decision per declared child.
 *
 * `Record`, never `Partial<Record>`: a partial map lets a new declared child
 * land without a decision, which is the drop this module exists to prevent.
 */
export type ChildHandlers<Container extends DispatchedContainer> = Readonly<
  Record<DeclaredChild<Container>, ChildDisposition>
>;

type DispatchChildrenOptions<Container extends DispatchedContainer> = {
  /** The container element, as the source wrote it. */
  element: XmlElement;
  /** Which container's declared-child set the handler map must be total over. */
  container: Container;
  handlers: ChildHandlers<Container>;
  /**
   * How many modelled children the caller holds right now. Called once per
   * captured child, so the capture lands after the siblings already read.
   */
  modelledCount: () => number;
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
  modelledCount,
  undeclared,
  undeclaredNamespaces,
}: DispatchChildrenOptions<Container>): PreservedMarkup | undefined => {
  const children: PreservedChild[] = [];
  const capture = (child: XmlElement): void => {
    children.push({ index: modelledCount(), xml: captureVerbatimXml(child) });
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
    if (disposition === DROPPED_WITH_ITS_WRAPPER || disposition === OWNED_ELSEWHERE) {
      continue;
    }
    disposition(child);
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
