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
 * reversed case is not decoration: without it a row whose order the validator
 * declines to check at all would pass the first assertion vacuously.
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
    const reversedVerdicts = members.map(([element, type]) => {
      const names = declaredByMember(element, type, CONTAINER_CHILDREN[key]);
      expect(names.length).toBeGreaterThan(1);
      expect(outOfOrderChildren(element, type, names)).toEqual([]);
      return outOfOrderChildren(element, type, [...names].reverse());
    });

    // The same names backwards. The validator only orders a content model it
    // cannot reorder itself, and `CT_SectPr` opens with a choice of header and
    // footer references, so it declines that one; `CT_SectPrBase` carries the
    // same sequence without them and is checked. One member that refuses the
    // reversal is what keeps the assertion above from passing vacuously.
    expect(reversedVerdicts.some((violations) => violations.length > 0)).toBe(true);
  });
});
