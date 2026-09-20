/**
 * A minimal schema-valid package that contains one pair under test.
 *
 * The survival law needs a document for every (container, allowed child) and
 * (element, allowed attribute) pair the schema declares inside the parts folio
 * rebuilds. Hand-writing them is not an option at this scale, so the fixture is
 * synthesised from the schema graph: the shortest chain of elements from the
 * part root down to the container, each level carrying the attributes its type
 * requires and the siblings its content model requires, and the subject placed
 * at the ordinal its particle declares.
 *
 * A fixture that does not itself validate is reported as *unrepresentable*
 * rather than run: a law that fails because the generator wrote invalid markup
 * proves nothing, and counting those separately is what keeps the census
 * honest about the containers it cannot reach.
 */

import {
  type ModelledAttributes,
  modelledAttributeNames,
  PROPERTY_ELEMENT_ATTRIBUTES,
} from "@stll/folio-core/docx/propertyElementAttributes";

import {
  type AttributeSlot,
  attributesOf,
  type ChildSlot,
  childrenOf,
  type Container,
  type ContainerId,
  containerKey,
  type ContainerSpace,
  qualify,
  type QualifiedName,
  type RebuiltPart,
  REBUILT_PARTS,
  type RebuiltPartRoot,
  type SchemaIndex,
  WML_NAMESPACE,
} from "./schemaSpace";
import { representativeValue } from "./values";

/** The part an element roots, when it roots one. */
const rebuiltPartOf = ({ namespace, name }: QualifiedName): RebuiltPart | undefined =>
  namespace === WML_NAMESPACE && name in REBUILT_PARTS
    ? // SAFETY: the `in` check proved the name a key of the table.
      REBUILT_PARTS[name as RebuiltPartRoot]
    : undefined;

const isRebuiltPartRoot = (element: QualifiedName): boolean => rebuiltPartOf(element) !== undefined;

const NAMESPACE_PREFIXES: ReadonlyArray<readonly [string, string]> = [
  ["w", WML_NAMESPACE],
  ["r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"],
  ["a", "http://schemas.openxmlformats.org/drawingml/2006/main"],
  ["wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"],
  ["pic", "http://schemas.openxmlformats.org/drawingml/2006/picture"],
  ["lc", "http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas"],
  ["m", "http://schemas.openxmlformats.org/officeDocument/2006/math"],
];

const PREFIX_BY_NAMESPACE = new Map(
  NAMESPACE_PREFIXES.map(([prefix, namespace]) => [namespace, prefix]),
);

export const XMLNS_DECLARATIONS = NAMESPACE_PREFIXES.map(
  ([prefix, namespace]) => ` xmlns:${prefix}="${namespace}"`,
).join("");

/** An attribute in no namespace is spelled without a prefix; that is not a missing value. */
export const spell = ({ namespace, name }: QualifiedName): string | undefined => {
  if (namespace === "") {
    return name;
  }
  const prefix = PREFIX_BY_NAMESPACE.get(namespace);
  return prefix === undefined ? undefined : `${prefix}:${name}`;
};

const DRAWINGML_MAIN_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/main";
const PICTURE_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/picture";

/**
 * The relationship the seeded `a:blip` binds to. The law's package carries a
 * 1×1 PNG under this id in every part it writes a fixture into.
 */
export const IMAGE_RELATIONSHIP_ID = "rIdContainerSurvivalImage";

/** How deep a required-sibling walk goes before it stops filling the content model. */
const REQUIRED_DEPTH_LIMIT = 3;

/** How many filler children one level generates before the fixture stops growing. */
const REQUIRED_SIBLING_LIMIT = 6;

/**
 * How many instances of one required particle the fixture writes.
 *
 * A particle declares how many times it has to appear, and `CT_WrapPath`
 * declares `minOccurs="2"` on `wp:lineTo`: a single `wp:lineTo` is markup no
 * conforming reader accepts, so a pair measured on it is measured on a
 * construct that cannot occur. The count is the declared minimum, capped by
 * {@link REQUIRED_SIBLING_LIMIT} so one particle can never be the thing that
 * makes a fixture unbounded. The cap is the same budget a level spends on
 * distinct fillers rather than a second number, and it does not bind today:
 * the largest `minOccurs` any particle in the graph declares is 3, so every
 * fixture sits exactly at the schema's minimum.
 *
 * The budget itself still counts particles rather than instances. It exists to
 * stop one level crowding out a sibling that makes the container what it is,
 * and repeating a particle the schema already demands takes nothing from any
 * other particle.
 */
const requiredInstances = (minOccurs: string): number => {
  const declared = Number.parseInt(minOccurs, 10);
  return Number.isInteger(declared) ? Math.min(declared, REQUIRED_SIBLING_LIMIT) : 1;
};

const escapeAttribute = (value: string): string =>
  value.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll('"', "&quot;");

export type WrittenAttribute = { spelled: string; value: string };

const writeAttributes = (attributes: readonly WrittenAttribute[]): string =>
  attributes.map(({ spelled, value }) => ` ${spelled}="${escapeAttribute(value)}"`).join("");

/** The attributes a type requires, each with the representative value of its simple type. */
const requiredAttributesOf = (
  index: SchemaIndex,
  typeQName: string | undefined,
  omit?: QualifiedName,
): WrittenAttribute[] => {
  if (typeQName === undefined) {
    return [];
  }
  const written: WrittenAttribute[] = [];
  const omitted = omit === undefined ? undefined : qualify(omit);
  for (const declared of attributesOf(index, `complexType:${typeQName}`)) {
    if (!declared.required || qualify(declared.attribute) === omitted) {
      continue;
    }
    const spelled = spell(declared.attribute);
    const value = declared.fixed ?? representativeValue(index, declared.typeQName);
    if (spelled !== undefined && value !== undefined) {
      written.push({ spelled, value });
    }
  }
  return written;
};

type Particle = {
  child: QualifiedName;
  typeQName: string | undefined;
  minOccurs: string;
  compositorId: string | undefined;
  order: number;
};

const particlesOf = (index: SchemaIndex, typeQName: string): Particle[] =>
  childrenOf(index, `complexType:${typeQName}`).map((child, order) => ({
    child: child.child,
    typeQName: child.typeQName,
    minOccurs: child.minOccurs,
    compositorId: child.compositorId,
    order,
  }));

const compositorKind = (index: SchemaIndex, compositorId: string | undefined): string =>
  compositorId === undefined ? "sequence" : (index.compositorKinds.get(compositorId) ?? "sequence");

/**
 * The children a content model requires, in declaration order.
 *
 * A particle inside a `choice` is only written when no member of that choice is
 * present yet, because writing two members of a choice is the kind of invalid
 * markup that would fail the law for the generator's reasons rather than the
 * subject's. Extra pieces (the subject, a `w:sectPr`) are merged in by ordinal.
 */
const renderChildren = (
  index: SchemaIndex,
  typeQName: string,
  depth: number,
  entered: ReadonlySet<string>,
  extra: ReadonlyArray<{ order: number; xml: string }> = [],
): string => {
  const pieces = [...extra];
  const satisfied = new Set<string>();
  for (const particle of extra.length === 0 ? [] : particlesOf(index, typeQName)) {
    if (extra.some(({ order }) => order === particle.order)) {
      satisfied.add(particle.compositorId ?? "");
    }
  }
  const nextEntered = new Set([...entered, typeQName]);
  // The budget is spent on fillers only. Counting the subject and the seeds
  // against it let a `wp:anchor` carrying either crowd out `a:graphic`, which
  // is the child that makes the drawing a drawing: the fixture then measured a
  // container the schema does not allow rather than the pair.
  let filled = 0;
  for (const particle of particlesOf(index, typeQName)) {
    if (particle.minOccurs === "0" || filled >= REQUIRED_SIBLING_LIMIT) {
      continue;
    }
    // The subject and the seeds are written at their particle's ordinal, and a
    // particle that declares more than one instance still owes the rest: the
    // fixture tops it up rather than skipping it.
    const present = pieces.filter(({ order }) => order === particle.order).length;
    const wanted = requiredInstances(particle.minOccurs);
    if (present >= wanted) {
      continue;
    }
    const compositor = particle.compositorId ?? "";
    // A choice takes one member. The exclusion is about which member, so it
    // only applies to a particle with nothing at its ordinal yet: one that is
    // being topped up is the member already chosen.
    if (present === 0 && compositorKind(index, particle.compositorId) === "choice") {
      if (satisfied.has(compositor)) {
        continue;
      }
      satisfied.add(compositor);
    }
    const rendered = renderFiller(index, particle.child, particle.typeQName, depth, nextEntered);
    if (rendered !== undefined) {
      pieces.push({ order: particle.order, xml: rendered.repeat(wanted - present) });
      filled += 1;
    }
  }
  return pieces
    .sort((left, right) => left.order - right.order)
    .map(({ xml }) => xml)
    .join("");
};

/**
 * A filler element that satisfies a content model without carrying a subject.
 *
 * It stops at {@link REQUIRED_DEPTH_LIMIT} and at any type it has already
 * entered: WordprocessingML's content models are mutually recursive (a table
 * cell holds a table), so an unguarded walk would not terminate.
 */
const renderFiller = (
  index: SchemaIndex,
  element: QualifiedName,
  typeQName: string | undefined,
  depth: number,
  entered: ReadonlySet<string>,
): string | undefined => {
  const spelled = spell(element);
  if (spelled === undefined) {
    return undefined;
  }
  const attributes = writeAttributes(requiredAttributesOf(index, typeQName));
  if (typeQName === undefined || depth >= REQUIRED_DEPTH_LIMIT || entered.has(typeQName)) {
    return `<${spelled}${attributes}/>`;
  }
  const inner = renderChildren(
    index,
    typeQName,
    depth + 1,
    entered,
    seedsFor(index, element.name, typeQName, []),
  );
  return inner === ""
    ? `<${spelled}${attributes}/>`
    : `<${spelled}${attributes}>${inner}</${spelled}>`;
};

/**
 * The partner a range marker needs to be a range.
 *
 * A lone `w:commentRangeStart` is not a document. The parse boundary now
 * tolerates one — it re-anchors an unmatched comment range as a point
 * `w:commentReference` and drops an unmatched move range, which is what Word
 * shows — so the fixture no longer needs the partner to get the package open.
 * It still needs it to measure anything: without a partner the census would
 * charge that deliberate normalisation to the pair and report a container as
 * losing a marker it never held. The partner carries the same `w:id` the
 * subject does, which is the representative value `ST_DecimalNumber` gets.
 *
 * What folio does with an orphan is a property of the normaliser and is tested
 * there (`inlineRangeMarkerTolerance.property.test.ts`); it is not what these
 * pairs are testing.
 */
const PARTNER_MARKERS: Readonly<Record<string, { xml: string; before: boolean }>> = {
  commentRangeStart: { xml: '<w:commentRangeEnd w:id="1"/>', before: false },
  commentRangeEnd: { xml: '<w:commentRangeStart w:id="1"/>', before: true },
  moveFromRangeStart: { xml: '<w:moveFromRangeEnd w:id="1"/>', before: false },
  // `CT_MoveBookmark` requires `w:author`; the corpus census found folio
  // writing these markers without it, so the partner must not repeat that.
  moveFromRangeEnd: {
    xml: '<w:moveFromRangeStart w:id="1" w:name="mv" w:author="folio1" w:date="2024-01-01T00:00:00Z"/>',
    before: true,
  },
  moveToRangeStart: { xml: '<w:moveToRangeEnd w:id="1"/>', before: false },
  moveToRangeEnd: {
    xml: '<w:moveToRangeStart w:id="1" w:name="mv" w:author="folio1" w:date="2024-01-01T00:00:00Z"/>',
    before: true,
  },
  bookmarkStart: { xml: '<w:bookmarkEnd w:id="1"/>', before: false },
  bookmarkEnd: { xml: '<w:bookmarkStart w:id="1" w:name="bm"/>', before: true },
};

type PartnerMarker = (typeof PARTNER_MARKERS)[string];

const withPartnerMarker = (childXml: string, partner: PartnerMarker | undefined): string => {
  if (partner === undefined) {
    return childXml;
  }
  return partner.before ? `${partner.xml}${childXml}` : `${childXml}${partner.xml}`;
};

export type Subject =
  | { kind: "child"; slot: ChildSlot }
  | {
      kind: "attribute";
      slot: AttributeSlot;
      value: string;
      /**
       * A second attribute to state on the same element, from
       * {@link modelledCompanionFor}.
       *
       * Without it the census measures one attribute at a time, and a reader
       * that decides an element whole passes every pair on an element it
       * models: `<w:ind w:leftChars="100"/>` is kept entire because the reader
       * took nothing from it, while the `<w:ind w:left="720"
       * w:leftChars="100"/>` a document actually carries loses the character
       * unit.
       */
      companion?: WrittenAttribute;
    };

export type BuiltFixture = {
  /** The package part this fixture is: `word/document.xml`, `word/styles.xml`, and so on. */
  part: RebuiltPart;
  /** The whole part the package carries at {@link BuiltFixture.part}'s path. */
  documentXml: string;
  /** The element the law looks for in the saved part, e.g. `w:tblGridChange`. */
  subjectSpelling: string;
  /** The same element qualified, for the equalities that depend on which slot it is. */
  subjectElement: QualifiedName;
  /** For an attribute subject, the attribute's spelling; its element is the container. */
  attributeSpelling: string | undefined;
  attributeLocalName: string | undefined;
};

export type FixtureResult =
  | { status: "built"; fixture: BuiltFixture }
  | { status: "unrepresentable"; reason: string };

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

/**
 * Content a container needs to survive at all, beyond what its content model requires.
 *
 * The schema lets a `w:tbl` hold no rows and a `w:tc` hold no paragraphs;
 * neither is a document any consumer keeps, and folio prunes both. A fixture
 * that is discarded for being empty would report every one of its pairs as
 * lost, which is a fact about the fixture rather than about folio. These seeds
 * are the minimum that makes the container real, keyed by its element name and
 * placed at the ordinal the schema gives the seeded child. They add content;
 * they never decide anything the contract decides.
 *
 * Each entry lists candidates and the first child the container's type actually
 * declares is the one written: `w:sdtContent` holds paragraphs under a block
 * content control, runs under an inline one and rows inside a table, and the
 * element name alone cannot tell the three apart.
 */
type Seed = {
  child: string;
  /** WordprocessingML unless the seeded child is declared elsewhere. */
  namespace?: string;
  xml: string;
};

const SEED_CHILDREN: Readonly<Record<string, ReadonlyArray<Seed>>> = {
  tbl: [{ child: "tr", xml: "<w:tr><w:tc><w:p/></w:tc></w:tr>" }],
  tr: [{ child: "tc", xml: "<w:tc><w:p/></w:tc>" }],
  tc: [{ child: "p", xml: "<w:p/>" }],
  p: [{ child: "r", xml: "<w:r><w:t>folio</w:t></w:r>" }],
  r: [{ child: "t", xml: "<w:t>folio</w:t>" }],
  hyperlink: [{ child: "r", xml: "<w:r><w:t>folio</w:t></w:r>" }],
  sdt: [{ child: "sdtContent", xml: "<w:sdtContent><w:p/></w:sdtContent>" }],
  sdtContent: [
    { child: "p", xml: "<w:p/>" },
    { child: "r", xml: "<w:r><w:t>folio</w:t></w:r>" },
    { child: "tr", xml: "<w:tr><w:tc><w:p/></w:tc></w:tr>" },
    { child: "tc", xml: "<w:tc><w:p/></w:tc>" },
  ],
  // A transparent inline wrapper says something about the content it holds and
  // nothing on its own, so the editor carries it as a mark on that content. An
  // empty `w:bdo` has no leaf to carry it and is dropped, which would read as
  // the wrapper being lost rather than as the fixture holding nothing.
  bdo: [{ child: "r", xml: "<w:r><w:t>folio</w:t></w:r>" }],
  dir: [{ child: "r", xml: "<w:r><w:t>folio</w:t></w:r>" }],
  // The same for the wrappers that name an element. `w:r` alone, not the
  // candidate list `w:sdtContent` needs: `w:customXml` is also a block, a row
  // and a cell wrapper, none of which declares a run, so naming only the run
  // leaves those three measured exactly as they were.
  smartTag: [{ child: "r", xml: "<w:r><w:t>folio</w:t></w:r>" }],
  customXml: [{ child: "r", xml: "<w:r><w:t>folio</w:t></w:r>" }],
  tblGrid: [{ child: "gridCol", xml: '<w:gridCol w:w="2400"/>' }],
  // A `w:numPr` that names no numbering is not a list, and folio drops it; the
  // `w:numberingChange` it can carry would then read as lost with it.
  numPr: [{ child: "ilvl", xml: '<w:ilvl w:val="0"/><w:numId w:val="1"/>' }],
  // A `w:drawing` with no `a:blip` names no picture relationship, and folio
  // classifies it preserve-only: the save replays its bytes whatever the
  // serializers would have written, so every pair under it read as surviving on
  // the strength of a byte copy and `serializeDrawingContent` was never
  // measured. Binding the blip to the package's synthetic PNG puts the drawing
  // on the rebuild path, which is the one an edited document takes.
  graphic: [
    {
      child: "graphicData",
      namespace: DRAWINGML_MAIN_NAMESPACE,
      xml:
        `<a:graphicData uri="${PICTURE_NAMESPACE}">` +
        '<pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="folio.png" descr="folio"/>' +
        "<pic:cNvPicPr/></pic:nvPicPr>" +
        `<pic:blipFill><a:blip r:embed="${IMAGE_RELATIONSHIP_ID}"/></pic:blipFill>` +
        "<pic:spPr/></pic:pic></a:graphicData>",
    },
  ],
};

/** One level of the chain: the element, its required attributes, its children in order. */
const renderLevel = (
  space: ContainerSpace,
  id: ContainerId,
  extra: ReadonlyArray<{ order: number; xml: string }>,
  extraAttributes: readonly WrittenAttribute[] = [],
  omitAttribute?: QualifiedName,
): string => {
  const spelled = spell(id.element);
  if (spelled === undefined) {
    return extra.map(({ xml }) => xml).join("");
  }
  const attributes = writeAttributes([
    ...requiredAttributesOf(space.index, id.typeQName, omitAttribute),
    ...extraAttributes,
  ]);
  // Every prefix is bound on the part's own root element: a part root is never
  // a child of anything, so membership in the table decides it without the
  // caller having to say which level it is rendering.
  const namespaces = isRebuiltPartRoot(id.element) ? XMLNS_DECLARATIONS : "";

  const seeded = [...extra, ...seedsFor(space.index, id.element.name, id.typeQName, extra)];

  // `w:body` without a `w:sectPr` is valid but nothing folio ever reads, and a
  // section it has to invent is a difference the law would charge to the pair.
  const withSection =
    id.element.name === "body" && !seeded.some(({ xml }) => xml.startsWith("<w:sectPr"))
      ? [...seeded, { order: Number.MAX_SAFE_INTEGER, xml: "<w:sectPr/>" }]
      : seeded;

  const body = renderChildren(space.index, id.typeQName, 1, new Set(), withSection);
  return body === ""
    ? `<${spelled}${namespaces}${attributes}/>`
    : `<${spelled}${namespaces}${attributes}>${body}</${spelled}>`;
};

const declaredOrdinal = (
  index: SchemaIndex,
  typeQName: string,
  child: QualifiedName,
): number | undefined => {
  const wanted = qualify(child);
  return particlesOf(index, typeQName).find((particle) => qualify(particle.child) === wanted)
    ?.order;
};

const ordinalOf = (index: SchemaIndex, typeQName: string, child: QualifiedName): number =>
  declaredOrdinal(index, typeQName, child) ?? 0;

/** The seeds a container needs, placed at the ordinals its content model gives them. */
const seedsFor = (
  index: SchemaIndex,
  elementName: string,
  typeQName: string,
  present: ReadonlyArray<{ order: number }>,
): Array<{ order: number; xml: string }> => {
  for (const seed of SEED_CHILDREN[elementName] ?? []) {
    const order = declaredOrdinal(index, typeQName, {
      namespace: seed.namespace ?? WML_NAMESPACE,
      name: seed.child,
    });
    if (order === undefined) {
      continue;
    }
    return present.some((piece) => piece.order === order) ? [] : [{ order, xml: seed.xml }];
  }
  return [];
};

/**
 * The chain from a rebuilt part's root down to the container, with the subject
 * at the bottom.
 *
 * Every root in {@link REBUILT_PARTS} is synthesised, so a container reachable
 * only under `w:styles`, `w:settings`, `w:hdr` or any other rebuilt part is
 * measured where it lives rather than reported as skipped. The law wraps the
 * result in a package that carries the part at its own path, with the
 * content-type override and the relationship that make a reader find it.
 */
export const buildFixture = (space: ContainerSpace, subject: Subject): FixtureResult => {
  const container = space.containers.get(containerKey(subject.slot.container));
  if (container === undefined) {
    return { status: "unrepresentable", reason: "container is not in the reachable space" };
  }
  const root = container.path.at(0);
  const part = root === undefined ? undefined : rebuiltPartOf(root.element);
  if (part === undefined) {
    return {
      status: "unrepresentable",
      reason: `reachable only under w:${root?.element.name ?? "?"}, which roots no rebuilt part`,
    };
  }

  const subjectSpelling = spell(
    subject.kind === "child" ? subject.slot.child : subject.slot.container.element,
  );
  const attributeSpelling =
    subject.kind === "attribute" ? spell(subject.slot.attribute) : undefined;
  if (
    subjectSpelling === undefined ||
    (subject.kind === "attribute" && attributeSpelling === undefined)
  ) {
    return { status: "unrepresentable", reason: "no prefix is bound for the subject's namespace" };
  }

  let xml = renderSubjectLevel(space, container, subject, attributeSpelling);
  if (xml === undefined) {
    return { status: "unrepresentable", reason: "no prefix is bound for the child's namespace" };
  }

  for (let level = container.path.length - 2; level >= 0; level -= 1) {
    // SAFETY: `level` and `level + 1` are indices the loop bound keeps inside the path.
    const ancestor = container.path[level] as ContainerId;
    // SAFETY: the same, one step deeper.
    const inner = container.path[level + 1] as ContainerId;
    xml = renderLevel(space, ancestor, [
      { order: ordinalOf(space.index, ancestor.typeQName, inner.element), xml },
    ]);
  }

  return {
    status: "built",
    fixture: {
      part,
      documentXml: `${XML_DECLARATION}${xml}`,
      subjectSpelling,
      subjectElement:
        subject.kind === "child" ? subject.slot.child : subject.slot.container.element,
      attributeSpelling,
      attributeLocalName: subject.kind === "attribute" ? subject.slot.attribute.name : undefined,
    },
  };
};

/** The property-element tables, widened to the key the census has in hand. */
const MODELLED_BY_TYPE: Readonly<Record<string, ModelledAttributes<string>>> =
  PROPERTY_ELEMENT_ATTRIBUTES;

/**
 * An attribute of the same element that the model has a field for, or nothing.
 *
 * Which attribute counts as modelled is the model's answer, not the census's:
 * `PROPERTY_ELEMENT_ATTRIBUTES` is the same table the reader computes its
 * remainder from, so the two cannot disagree about what "modelled" means. The
 * The first field other than the subject's own whose declared attribute has a
 * representative value is the one written, which makes the choice
 * deterministic across runs without pairing alternate spellings of one field.
 *
 * Nothing is returned when the element declares a *required* modelled
 * attribute: the fixture already states every required attribute, so the
 * subject is beside a modelled sibling in the ordinary fixture and a second
 * one would measure the same package twice. `CT_TabStop` is that case.
 */
export const modelledCompanionFor = (
  space: ContainerSpace,
  slot: AttributeSlot,
): WrittenAttribute | undefined => {
  const modelled = MODELLED_BY_TYPE[localTypeName(slot.container.typeQName)];
  const container = space.containers.get(containerKey(slot.container));
  if (modelled === undefined || container === undefined) {
    return undefined;
  }

  const fields = Object.values(modelled).map((declared) =>
    typeof declared === "string" ? [declared] : Array.from(declared),
  );
  for (const name of modelledAttributeNames(modelled)) {
    const declared = container.attributes.find(
      ({ attribute }) => attribute.namespace === WML_NAMESPACE && attribute.name === name,
    );
    if (declared?.required) {
      return undefined;
    }
  }

  for (const field of fields) {
    // Two names in one field are alternate spellings or mutually exclusive
    // encodings of the same value. Stating both would manufacture a conflict
    // rather than place the subject beside another modelled attribute.
    if (field.includes(slot.attribute.name)) {
      continue;
    }
    for (const name of field) {
      const declared = container.attributes.find(
        ({ attribute }) => attribute.namespace === WML_NAMESPACE && attribute.name === name,
      );
      if (declared === undefined) {
        continue;
      }
      const spelled = spell(declared.attribute);
      const value = declared.fixed ?? representativeValue(space.index, declared.typeQName);
      if (spelled !== undefined && value !== undefined) {
        return { spelled, value };
      }
    }
  }
  return undefined;
};

const localTypeName = (typeQName: string): string => typeQName.slice(typeQName.indexOf("}") + 1);

const renderSubjectLevel = (
  space: ContainerSpace,
  container: Container,
  subject: Subject,
  attributeSpelling: string | undefined,
): string | undefined => {
  if (subject.kind === "attribute") {
    return renderLevel(
      space,
      container.id,
      [],
      [
        { spelled: attributeSpelling ?? "", value: subject.value },
        ...(subject.companion === undefined ? [] : [subject.companion]),
      ],
      subject.slot.attribute,
    );
  }
  const childXml = renderFiller(
    space.index,
    subject.slot.child,
    subject.slot.childTypeQName,
    1,
    new Set([container.id.typeQName]),
  );
  if (childXml === undefined) {
    return undefined;
  }
  const order = ordinalOf(space.index, container.id.typeQName, subject.slot.child);
  const partner =
    subject.slot.child.namespace === WML_NAMESPACE
      ? PARTNER_MARKERS[subject.slot.child.name]
      : undefined;
  const paired = withPartnerMarker(childXml, partner);
  return renderLevel(space, container.id, [{ order, xml: paired }]);
};
