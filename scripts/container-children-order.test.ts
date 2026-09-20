/**
 * A sequence row's generated order is the corpus validator's order.
 *
 * Three readers derive an ordinal from the same schema particles:
 * `container-survival/schemaSpace.ts` builds a fixture with its subject at one,
 * `corpus-schema-validator.ts` scores a document's children against one, and
 * `generate-container-children.ts` writes one into `CONTAINER_CHILDREN` for
 * every `sequence: true` row — which is the order `serializeSequenceChildren`
 * writes a property set in. They share one derivation
 * (`orderedParticlesByOwner`), and this is the check that they still agree
 * about the result: a serializer whose order is not the validator's writes a
 * property set a consumer refuses.
 *
 * The question is asked of the validator's public surface rather than of its
 * ordinal table, so the test measures the verdict a corpus run would get. The
 * reversed case is not decoration: a clean verdict means both "in order" and
 * "there is no order to be in", and without the reversal a member the validator
 * declines to score would pass the first assertion vacuously.
 *
 * Which members it declines is not a fact about folio and not something to
 * exempt: `contentModelFor` refuses an order wherever the model can reorder
 * itself, and `CT_TcPr` reaches `EG_CellMarkupElements` and `CT_TrPrBase` is
 * itself a repeated choice. So `ordersChildrenOf` is asked first and every
 * member is then pinned to a definite verdict — a scored one must refuse the
 * reversal, an unscored one must stay silent in both directions. A validator
 * that quietly stops ordering `CT_TblPrBase`, or starts ordering `CT_TcPr`
 * without the generator agreeing, fails here rather than passing.
 *
 * A row no member scores keeps its serializer honest elsewhere:
 * `tableCellPropertySet.property.test.ts` writes every declared child of a
 * `w:tcPr` through the serializer and reads the order back off the schema.
 */

import { describe, expect, test } from "bun:test";

import {
  CONTAINER_CHILDREN,
  SEQUENCE_CONTAINERS,
} from "../packages/core/src/docx/containerChildren.gen";
import { DISPATCHED_CONTAINERS } from "./lib/container-survival/dispatchedContainers";
import {
  containerKey,
  loadContainerSpace,
  qualify,
  WML_NAMESPACE,
} from "./lib/container-survival/schemaSpace";
import {
  loadSchemaGraph,
  ordersChildrenOf,
  SCHEMA_VIOLATION_KINDS,
  validateOoxmlPart,
} from "./lib/corpus-schema-validator";

const space = await loadContainerSpace();
const graph = await loadSchemaGraph();

/**
 * The container nested at the chain the census walks to reach it, holding the
 * children named.
 *
 * Only the order check is read off the result, so the ancestors carry no
 * attributes and the children are empty elements: a missing required attribute
 * is a violation of a different kind at a different path, and one this test has
 * no opinion about.
 */
const partWithChildren = (element: string, type: string, children: readonly string[]): string => {
  const container = space.containers.get(
    containerKey({
      element: { namespace: WML_NAMESPACE, name: element },
      typeQName: qualify({ namespace: WML_NAMESPACE, name: type }),
    }),
  );
  if (!container) {
    throw new Error(`the schema graph has no container ${element}|${type}`);
  }
  const chain = container.path.map(({ element: { name } }) => name);
  const inner = children.map((name) => `<w:${name}/>`).join("");
  const open = chain.map((name) => `<w:${name}>`).join("");
  const close = [...chain]
    .reverse()
    .map((name) => `</w:${name}>`)
    .join("");
  return `<?xml version="1.0"?>${open.replace(
    "<w:document>",
    `<w:document xmlns:w="${WML_NAMESPACE}">`,
  )}${inner}${close}`;
};

const outOfOrderChildren = (element: string, type: string, children: readonly string[]): string[] =>
  validateOoxmlPart({ graph, xml: partWithChildren(element, type, children) })
    .filter(({ kind }) => kind === SCHEMA_VIOLATION_KINDS.outOfOrderChild)
    .map(({ name }) => name);

/** Only the names the member itself declares; `CT_TblPrBase` has no change. */
const declaredByMember = (element: string, type: string, names: readonly string[]): string[] => {
  const container = space.containers.get(
    containerKey({
      element: { namespace: WML_NAMESPACE, name: element },
      typeQName: qualify({ namespace: WML_NAMESPACE, name: type }),
    }),
  );
  const own = new Set(
    (container?.children ?? [])
      .filter(({ child }) => child.namespace === WML_NAMESPACE)
      .map(({ child }) => child.name),
  );
  return names.filter((name) => own.has(name));
};

const sequenceRows = SEQUENCE_CONTAINERS.map((key) => {
  const row = DISPATCHED_CONTAINERS.find((candidate) => candidate.key === key);
  if (!row) {
    throw new Error(`no dispatched container is keyed ${key}`);
  }
  return { key, members: row.members };
});

describe("a generated sequence is the order the corpus validator scores against", () => {
  test.each(sequenceRows)("$key", ({ key, members }) => {
    for (const [element, type] of members) {
      const names = declaredByMember(element, type, CONTAINER_CHILDREN[key]);
      expect(names.length).toBeGreaterThan(1);
      expect(outOfOrderChildren(element, type, names)).toEqual([]);

      // The same names backwards, against the validator's own answer about
      // whether it has an order for this type. `CT_SectPrBase` is scored and
      // `CT_SectPr` is not — it opens with a choice of header and footer
      // references — so the two verdicts differ for one row, and pinning each
      // member to the one it owes is what keeps the assertion above from
      // passing vacuously.
      const reversed = outOfOrderChildren(element, type, [...names].reverse());
      if (ordersChildrenOf(graph, qualify({ namespace: WML_NAMESPACE, name: type }))) {
        expect(reversed.length).toBeGreaterThan(0);
        continue;
      }
      expect(reversed).toEqual([]);
    }
  });
});
