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
  type ContainerSpace,
  loadContainerSpace,
  qualify,
  WML_NAMESPACE,
} from "./lib/container-survival/schemaSpace";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const OUTPUT_PATH = path.join(REPO_ROOT, "packages/core/src/docx/containerChildren.gen.ts");
/**
 * The sequence rows are emitted here instead, and folio-core's table spreads
 * them in.
 *
 * Two packages write a property set — `@stll/docx-core`'s build-from-scratch
 * export writes `w:rPr`, and folio-core's serializers write all three — and the
 * dependency runs one way, so a table in folio-core is a table docx-core cannot
 * read. Emitting the order into the lower package and deriving the upper one's
 * table from it leaves one order for both.
 */
const SEQUENCE_OUTPUT_PATH = path.join(
  REPO_ROOT,
  "packages/docx-core/src/schema/sequenceChildren.gen.ts",
);

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
type DispatchedContainer = {
  key: string;
  members: readonly (readonly [element: string, type: string])[];
  /**
   * The members' children have one declared order, and the generated list is
   * written in it rather than sorted.
   *
   * A property set is the case. How binding the order is differs by container
   * and the difference is worth stating: `CT_TblPrBase` is a real
   * `xsd:sequence` of distinct names and a validating consumer refuses a
   * `w:tblPr` written in any other order, while `EG_RPrBase` is an
   * `xsd:choice` referenced `maxOccurs="unbounded"`, so a `w:rPr` in any order
   * is valid and folio's order is a canonical form rather than a requirement.
   * Either way there is one order, it comes from the schema's declaration
   * order, and it is the generated list itself rather than a second table
   * beside it — so a serializer that sorts by it cannot drift from the set the
   * handler map is total over.
   */
  sequence?: true;
};

const DISPATCHED_CONTAINERS: readonly DispatchedContainer[] = [
  { key: "w:comment", members: [["comment", "CT_Comment"]] },
  {
    // One walk serves a paragraph, a run-level tracked-change wrapper, a
    // bidirectional wrapper, a smart tag and an inline content control: they
    // share `EG_PContent`/`EG_ContentRunContent` and folio reads them with one
    // function, so one map has to be total over everything any of them holds.
    key: "run-level-content",
    members: [
      ["p", "CT_P"],
      ["ins", "CT_RunTrackChange"],
      ["smartTag", "CT_SmartTagRun"],
      ["bdo", "CT_BdoContentRun"],
      ["dir", "CT_DirContentRun"],
      ["sdtContent", "CT_SdtContentRun"],
    ],
  },
  {
    // A table and a row-level content control share one walk: folio unwraps
    // the control and splices its rows into the table. A `w:customXml` row
    // wrapper is kept whole rather than unwrapped, so this walk never descends
    // into `CT_CustomXmlRow`; it is a member anyway, because the union is the
    // safe direction and its one extra name — `w:customXmlPr` — then carries a
    // decision instead of falling to a default.
    key: "table-content",
    members: [
      ["tbl", "CT_Tbl"],
      ["sdtContent", "CT_SdtContentRow"],
      ["customXml", "CT_CustomXmlRow"],
    ],
  },
  {
    // A row and a row-level content control share one walk: folio unwraps the
    // control and splices its rows' content into the row, so one map has to be
    // total over everything either may hold.
    key: "row-content",
    members: [
      ["tr", "CT_Row"],
      ["sdtContent", "CT_SdtContentRow"],
    ],
  },
  // A link and a simple field each hold their own subset of `EG_PContent`
  // and each has its own parser, so each gets its own row rather than
  // borrowing `run-level-content`: the union would make a handler map total
  // over names the container cannot hold, and the two parsers would then
  // record decisions for children that never reach them.
  { key: "w:hyperlink", members: [["hyperlink", "CT_Hyperlink"]] },
  { key: "w:fldSimple", members: [["fldSimple", "CT_SimpleField"]] },
  {
    key: "block-content",
    members: [
      ["body", "CT_Body"],
      ["hdr", "CT_HdrFtr"],
      ["ftr", "CT_HdrFtr"],
      ["tc", "CT_Tc"],
      ["sdtContent", "CT_SdtContentBlock"],
      ["footnote", "CT_FtnEdn"],
      ["endnote", "CT_FtnEdn"],
    ],
  },
  {
    // A table's own property set, and the one a style carries. `CT_TblPr` is
    // `CT_TblPrBase` plus `w:tblPrChange`, so the two agree on every child
    // they share and the merged sequence is the wider of the two.
    key: "table-properties",
    members: [
      ["tblPr", "CT_TblPr"],
      ["tblPr", "CT_TblPrBase"],
    ],
    sequence: true,
  },
  {
    // One property set with four owners: a run's own `w:rPr`, the paragraph
    // mark's, and the snapshot each of their `w:rPrChange` records holds.
    // `CT_RPr` is `EG_RPrBase` plus `w:rPrChange`; `CT_ParaRPr` opens with
    // `EG_ParaRPrTrackChanges` on top of that, so the merged sequence is the
    // widest of the four and every owner's children carry one decision.
    key: "run-properties",
    members: [
      ["rPr", "CT_ParaRPr"],
      ["rPr", "CT_RPr"],
      ["rPr", "CT_ParaRPrOriginal"],
      ["rPr", "CT_RPrOriginal"],
    ],
    sequence: true,
  },
  {
    // A section's properties, and the snapshot a `w:sectPrChange` holds.
    // `CT_SectPrBase` is `CT_SectPr` without the two header/footer references
    // that open it and without the change that closes it.
    key: "section-properties",
    members: [
      ["sectPr", "CT_SectPr"],
      ["sectPr", "CT_SectPrBase"],
    ],
    sequence: true,
  },
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

const sequenceHeader = `/**
 * GENERATED FILE — do not edit.
 *
 * The containers whose content model declares one order for its children, with
 * the children written in it, derived from the committed schema graph by
 * \`scripts/generate-container-children.ts\`.
 *
 * It lives in the lower package because two packages write a property set and
 * the dependency runs one way: \`@stll/docx-core\` compiles a \`w:rPr\` from a
 * legal source and cannot import \`@stll/folio-core\`. folio-core's
 * declared-child table spreads this one in, so both write the same order and
 * neither restates it.
 *
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

const row = (key: string, names: readonly string[]): string =>
  `  ${JSON.stringify(key)}: [${names.map((name) => JSON.stringify(name)).join(", ")}],`;

type GeneratedFile = { path: string; contents: string };

const render = async (): Promise<GeneratedFile[]> => {
  const space = await loadContainerSpace();
  const setRows: string[] = [];
  const sequenceRows: string[] = [];

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
    (sequence ? sequenceRows : setRows).push(row(key, names));
  }

  const sequenceFile = [
    sequenceHeader,
    "export const SEQUENCE_CHILDREN = {",
    ...sequenceRows,
    "} as const;",
    "",
    "/** A container whose declared children have an order. */",
    "export type SequenceContainer = keyof typeof SEQUENCE_CHILDREN;",
    "",
    "/** Every child the schema declares for `Container`, in declaration order. */",
    "export type SequenceChild<Container extends SequenceContainer> =",
    "  (typeof SEQUENCE_CHILDREN)[Container][number];",
    "",
  ].join("\n");

  const childrenFile = [
    header,
    'import { SEQUENCE_CHILDREN } from "@stll/docx-core/schema";',
    "",
    "export const CONTAINER_CHILDREN = {",
    ...setRows,
    "  // Written in schema order and owned by `@stll/docx-core`, because a",
    "  // serializer there writes one of them too; see its module comment.",
    "  ...SEQUENCE_CHILDREN,",
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

  return [
    { path: SEQUENCE_OUTPUT_PATH, contents: sequenceFile },
    { path: OUTPUT_PATH, contents: childrenFile },
  ];
};

const mode = Bun.argv[2];
const generated = await render();

if (mode === "write") {
  for (const file of generated) {
    await writeFile(file.path, file.contents, "utf8");
    console.log(`wrote ${path.relative(REPO_ROOT, file.path)}`);
  }
} else if (mode === "check") {
  for (const file of generated) {
    const current = await Bun.file(file.path)
      .text()
      .catch(() => "");
    if (current !== file.contents) {
      throw new GenerateContainerChildrenError({
        message:
          `${path.relative(REPO_ROOT, file.path)} is stale. ` +
          "Run `bun run generate:container-children`.",
      });
    }
    console.log(`${path.relative(REPO_ROOT, file.path)} is up to date`);
  }
} else {
  throw new GenerateContainerChildrenError({
    message: "Usage: bun scripts/generate-container-children.ts <write|check>",
  });
}
