/**
 * Generate the declared-child name list for every container folio dispatches.
 *
 * The shared child dispatcher's handler map has to be **total** over the
 * children a container's content model declares, or a container can gain a
 * declared child without anybody deciding what happens to it — the exact
 * omission the container contract exists to stop. Totality wants a key union,
 * and a key union wants a source of truth that is not a hand-written list.
 *
 * `docs/container-contract.md` explains why the contract itself keeps its keys
 * as strings: a union over three thousand schema pairs would cost more type
 * instantiations than every published package put together. That argument does
 * not apply here. A container declares a few dozen children, the dispatched
 * set is small, and a union of that size is free — so the names are generated
 * into a committed `as const` table and the handler maps are checked against
 * it by the compiler.
 *
 * The list is derived from the same container space the survival census walks,
 * so the dispatcher, the census and the contract all read one schema.
 *
 * Usage:
 *   bun scripts/generate-container-children.ts write
 *   bun scripts/generate-container-children.ts check
 */

import { writeFile } from "node:fs/promises";
import path from "node:path";

import { TaggedError } from "better-result";

import { DISPATCHED_CONTAINERS } from "./lib/container-survival/dispatchedContainers";
import {
  containerKey,
  type ContainerSpace,
  loadContainerSpace,
  qualify,
  WML_NAMESPACE,
} from "./lib/container-survival/schemaSpace";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const OUTPUT_PATH = path.join(REPO_ROOT, "packages/core/src/docx/containerChildren.gen.ts");

class GenerateContainerChildrenError extends TaggedError("GenerateContainerChildrenError")<{
  message: string;
}> {}

const header = `/**
 * GENERATED FILE — do not edit.
 *
 * Every child the Transitional content model declares for a container folio
 * dispatches, derived from the committed schema graph by
 * \`scripts/generate-container-children.ts\`. The shared child dispatcher makes
 * a handler map total over these names, so a container cannot gain a declared
 * child without somebody deciding whether it is modelled or captured.
 * Regenerate with:
 *
 *   bun run generate:container-children
 */
`;

/**
 * One member's children, in the order its content model declares them.
 *
 * `loadContainerSpace` yields a container's children in declaration order, so
 * the sequence is the list itself; the check the generator makes is that two
 * members never disagree about it.
 */
const declaredChildren = (space: ContainerSpace, element: string, type: string): string[] => {
  const memberKey = containerKey({
    element: { namespace: WML_NAMESPACE, name: element },
    typeQName: qualify({ namespace: WML_NAMESPACE, name: type }),
  });
  const container = space.containers.get(memberKey);
  if (!container) {
    throw new GenerateContainerChildrenError({
      message: `the schema graph has no container ${memberKey}`,
    });
  }
  return container.children
    .filter(({ child }) => child.namespace === WML_NAMESPACE)
    .map(({ child }) => child.name);
};

/**
 * The members' sequences merged into one, or a refusal when they disagree.
 *
 * Two members of a sequence row describe the same property set at different
 * widths — `CT_TblPrBase` is `CT_TblPr` without the change record — so one
 * sequence contains the other and merging is inserting the missing names at
 * the position the wider member gives them. Two members that order a shared
 * child differently have no single sequence, and a serializer sorting by a
 * merged one would write markup a consumer refuses; that is a generator
 * failure rather than a choice to make here.
 */
const mergedSequence = (key: string, sequences: readonly (readonly string[])[]): string[] => {
  const merged = [...(sequences.at(0) ?? [])];
  for (const sequence of sequences.slice(1)) {
    let at = 0;
    for (const name of sequence) {
      const found = merged.indexOf(name);
      if (found === -1) {
        merged.splice(at, 0, name);
        at += 1;
        continue;
      }
      if (found < at) {
        throw new GenerateContainerChildrenError({
          message: `container ${key} has members that declare w:${name} in different orders`,
        });
      }
      at = found + 1;
    }
  }
  return merged;
};

const render = async (): Promise<string> => {
  const space = await loadContainerSpace();
  const rows: string[] = [];
  const sequences: string[] = [];

  for (const { key, members, sequence } of DISPATCHED_CONTAINERS) {
    const perMember = members.map(([element, type]) => declaredChildren(space, element, type));
    const names = sequence
      ? mergedSequence(key, perMember)
      : [...new Set(perMember.flat())].toSorted();
    if (names.length === 0) {
      throw new GenerateContainerChildrenError({
        message: `container ${key} declares no wordprocessingml children`,
      });
    }
    rows.push(
      `  ${JSON.stringify(key)}: [${names.map((name) => JSON.stringify(name)).join(", ")}],`,
    );
    if (sequence) {
      sequences.push(`  ${JSON.stringify(key)},`);
    }
  }

  return [
    header,
    "export const CONTAINER_CHILDREN = {",
    ...rows,
    "} as const;",
    "",
    "/** A container the shared child dispatcher covers. */",
    "export type DispatchedContainer = keyof typeof CONTAINER_CHILDREN;",
    "",
    "/** Every child name the schema declares for `Container`. */",
    "export type DeclaredChild<Container extends DispatchedContainer> =",
    "  (typeof CONTAINER_CHILDREN)[Container][number];",
    "",
    "/**",
    " * The containers whose content model is one flat sequence.",
    " *",
    " * Their entry in {@link CONTAINER_CHILDREN} is written in the order the",
    " * schema declares rather than sorted, so a child's position in that list is",
    " * its position in the element a serializer writes. Every other container's",
    " * list is a set and says nothing about order.",
    " */",
    `export const SEQUENCE_CONTAINERS = [`,
    ...sequences,
    "] as const;",
    "",
    "/** A container whose declared children have an order. */",
    "export type SequenceContainer = (typeof SEQUENCE_CONTAINERS)[number];",
    "",
  ].join("\n");
};

const mode = Bun.argv[2];
const generated = await render();

if (mode === "write") {
  await writeFile(OUTPUT_PATH, generated, "utf8");
  console.log(`wrote ${path.relative(REPO_ROOT, OUTPUT_PATH)}`);
} else if (mode === "check") {
  const current = await Bun.file(OUTPUT_PATH)
    .text()
    .catch(() => "");
  if (current !== generated) {
    throw new GenerateContainerChildrenError({
      message:
        `${path.relative(REPO_ROOT, OUTPUT_PATH)} is stale. ` +
        "Run `bun run generate:container-children`.",
    });
  }
  console.log(`${path.relative(REPO_ROOT, OUTPUT_PATH)} is up to date`);
} else {
  throw new GenerateContainerChildrenError({
    message: "Usage: bun scripts/generate-container-children.ts <write|check>",
  });
}
