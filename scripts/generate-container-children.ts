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

import {
  containerKey,
  loadContainerSpace,
  qualify,
  WML_NAMESPACE,
} from "./lib/container-survival/schemaSpace";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const OUTPUT_PATH = path.join(REPO_ROOT, "packages/core/src/docx/containerChildren.gen.ts");

class GenerateContainerChildrenError extends TaggedError("GenerateContainerChildrenError")<{
  message: string;
}> {}

/**
 * The containers routed through the shared dispatcher.
 *
 * A row is one handler map. Its members are the `w:`-local element names and
 * the complex types the census charges their pairs to, and the generated name
 * list is their union: one walker serves `w:body`, a header, a cell, an SDT's
 * content and a note, so one map has to be total over everything any of them
 * may hold. A union over-approximates for each member alone, which is the safe
 * direction — a decision recorded for a child a given container cannot hold
 * costs a line; a child with no decision costs the markup.
 *
 * A container joins this table when its parser is migrated. Adding a row is
 * how a migration declares itself: the generated union then makes the new
 * handler map's gaps a compile error rather than a silent drop.
 */
const DISPATCHED_CONTAINERS: readonly (readonly [
  key: string,
  members: readonly (readonly [element: string, type: string])[],
])[] = [
  ["w:comment", [["comment", "CT_Comment"]]],
  [
    // One walk serves a paragraph, a run-level tracked-change wrapper, a
    // bidirectional wrapper, a smart tag and an inline content control: they
    // share `EG_PContent`/`EG_ContentRunContent` and folio reads them with one
    // function, so one map has to be total over everything any of them holds.
    "run-level-content",
    [
      ["p", "CT_P"],
      ["ins", "CT_RunTrackChange"],
      ["smartTag", "CT_SmartTagRun"],
      ["bdo", "CT_BdoContentRun"],
      ["dir", "CT_DirContentRun"],
      ["sdtContent", "CT_SdtContentRun"],
    ],
  ],
  [
    "block-content",
    [
      ["body", "CT_Body"],
      ["hdr", "CT_HdrFtr"],
      ["ftr", "CT_HdrFtr"],
      ["tc", "CT_Tc"],
      ["sdtContent", "CT_SdtContentBlock"],
      ["footnote", "CT_FtnEdn"],
      ["endnote", "CT_FtnEdn"],
    ],
  ],
];

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

const render = async (): Promise<string> => {
  const space = await loadContainerSpace();
  const rows: string[] = [];

  for (const [key, members] of DISPATCHED_CONTAINERS) {
    const names = new Set<string>();
    for (const [element, type] of members) {
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
      for (const { child } of container.children) {
        if (child.namespace === WML_NAMESPACE) {
          names.add(child.name);
        }
      }
    }
    if (names.size === 0) {
      throw new GenerateContainerChildrenError({
        message: `container ${key} declares no wordprocessingml children`,
      });
    }
    rows.push(
      `  ${JSON.stringify(key)}: [${[...names]
        .toSorted()
        .map((name) => JSON.stringify(name))
        .join(", ")}],`,
    );
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
