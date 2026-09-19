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

import { escapeXmlAttribute } from "@stll/docx-core";
import type { PreservedAttribute, PreservedChild, PreservedMarkup } from "@stll/docx-core/model";

import type { DeclaredChild, DispatchedContainer } from "./containerChildren.gen";
import { captureVerbatimXml } from "./verbatimCapture";
import { getAttributes, getChildElements, getLocalName, type XmlElement } from "./xmlParser";

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

/** What a container does with one declared child. */
export type ChildDisposition =
  | ((child: XmlElement) => void)
  | typeof CAPTURE
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
   * Whether the caller's model already holds this attribute, by qualified
   * name as the source spells it. Everything it declines is kept in the
   * ordered attribute remainder.
   *
   * A predicate rather than a name set because a prefix is not an identity:
   * folio resolves an attribute by namespace URI plus local name, and a
   * remainder built by matching `"w:id"` textually would keep a second copy
   * of a `w:id` a source spelled `altw:id`. Omitting it leaves attributes
   * alone entirely, which is the safe default for a container whose parser
   * has not yet been asked the question.
   */
  modelsAttribute?: (name: string) => boolean;
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
};

const isNamespaceDeclaration = (name: string): boolean =>
  name === "xmlns" || name.startsWith("xmlns:");

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
  modelsAttribute,
  undeclared,
}: DispatchChildrenOptions<Container>): PreservedMarkup | undefined => {
  const children: PreservedChild[] = [];
  const capture = (child: XmlElement): void => {
    children.push({ index: modelledCount(), xml: captureVerbatimXml(child) });
  };

  // A declared name wins, so a name that somehow appears in both is decided
  // by the map the compiler checked.
  const dispositionByName = new Map<string, ChildDisposition>(Object.entries(handlers));
  for (const [name, disposition] of Object.entries(undeclared ?? {})) {
    if (!dispositionByName.has(name)) {
      dispositionByName.set(name, disposition);
    }
  }
  for (const child of getChildElements(element)) {
    // A local name is looked up against the declared set, not the qualified
    // name: the schema declares these in one namespace, and a child in any
    // other is undeclared by construction and belongs in the sink.
    const disposition = dispositionByName.get(getLocalName(child.name));
    if (disposition === undefined || disposition === CAPTURE) {
      capture(child);
      continue;
    }
    if (disposition === OWNED_ELSEWHERE) {
      continue;
    }
    disposition(child);
  }

  const attributes: PreservedAttribute[] = [];
  if (modelsAttribute) {
    for (const [name, value] of Object.entries(getAttributes(element))) {
      // A namespace declaration is not content: `captureVerbatimXml` rebinds
      // what a captured fragment needs, and replaying the container's own
      // bindings onto a rebuilt root would fight the root's.
      if (isNamespaceDeclaration(name) || modelsAttribute(name)) {
        continue;
      }
      attributes.push({ name, value });
    }
  }

  if (children.length === 0 && attributes.length === 0) {
    return undefined;
  }
  return {
    ...(children.length > 0 ? { children } : {}),
    ...(attributes.length > 0 ? { attributes } : {}),
  };
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

/** The container's preserved attributes, ready to append to its start tag. */
export const serializePreservedAttributes = (preserved: PreservedMarkup | undefined): string =>
  (preserved?.attributes ?? [])
    .map(({ name, value }) => ` ${name}="${escapeXmlAttribute(value)}"`)
    .join("");
