/**
 * Generate the model enumerations that mirror an OOXML simple type.
 *
 * A hand-written union that stands for an enumeration drifts silently. The
 * reader narrows an attribute against the union, `narrowEnum` returns
 * `undefined` for a token the union omits, and the caller drops the attribute:
 * the value never reaches the writer, and nothing in the build compares the two
 * lists. `ST_ThemeColor` drifted that way for sixteen members; these had
 * drifted too, each by a handful of members the format declares and folio did
 * not spell.
 *
 * So the member lists come from the committed schema graph:
 *
 *   ParagraphAlignment   w:ST_Jc              `w:jc/@w:val`
 *   TabStopAlignment     w:ST_TabJc           `w:tab/@w:val`
 *   TableAlignment       w:ST_JcTable         `w:tblPr/w:jc/@w:val`
 *   TextDirection        w:ST_TextDirection   `w:textDirection/@w:val`
 *   NumberFormat         w:ST_NumberFormat    `w:numFmt/@w:val`
 *
 * `ST_TextDirection` also gets its flow map, because the enumeration holds two
 * spellings of each of six flows and nothing in the schema says which pairs
 * with which. {@link TEXT_DIRECTION_FLOW_ENTRIES} carries that pairing with its
 * citation, and the checks below hold it to the enumeration.
 *
 * Usage:
 *   bun scripts/generate-ooxml-enumerations.ts write
 *   bun scripts/generate-ooxml-enumerations.ts check
 */

import path from "node:path";

import { TaggedError } from "better-result";

import { buildIndex, loadSchemaGraph, WML_NAMESPACE } from "./lib/ooxml-schema-graph";
import {
  emitGeneratedModule,
  enumerationOf,
  renderList,
  renderMap,
  renderModule,
} from "./lib/generated-enumeration";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const OUTPUT_PATH = path.join(REPO_ROOT, "packages/docx-core/src/model/ooxmlEnumerations.gen.ts");
const SCRIPT = "generate:ooxml-enumerations";

class TextDirectionPairingError extends TaggedError("TextDirectionPairingError")<{
  message: string;
}> {}

/**
 * Each `ST_TextDirection` token and the flow it names.
 *
 * ECMA-376 Part 4 §14.11.7 lists the six tokens Transitional adds over Strict
 * and gives each one as semantically equivalent to a Strict token: `btLr` to
 * `lr`, `lrTb` to `tb`, `lrTbV` to `tbV`, `tbLrV` to `lrV`, `tbRl` to `rl` and
 * `tbRlV` to `rlV`. The Strict spelling is the flow's name here, so the six
 * flows are exactly Strict's `ST_TextDirection` (Part 1 §17.18.93) and a
 * Transitional document and its Strict twin reach the same flow.
 *
 * The pairing is prose, not schema, so it is written out; the checks below are
 * what keep it from drifting from the enumeration it partitions.
 */
const TEXT_DIRECTION_FLOW_ENTRIES: readonly (readonly [string, string])[] = [
  ["tb", "tb"],
  ["rl", "rl"],
  ["lr", "lr"],
  ["tbV", "tbV"],
  ["rlV", "rlV"],
  ["lrV", "lrV"],
  ["btLr", "lr"],
  ["lrTb", "tb"],
  ["lrTbV", "tbV"],
  ["tbLrV", "lrV"],
  ["tbRl", "rl"],
  ["tbRlV", "rlV"],
];

/**
 * Hold the pairing to the enumeration it partitions.
 *
 * Every token is paired exactly once, every flow is itself a token, and a flow
 * pairs with itself: that last one is what makes the flow names Strict's own
 * spellings rather than a vocabulary of folio's.
 */
const textDirectionFlows = (tokens: readonly string[]): readonly string[] => {
  const paired = new Map(TEXT_DIRECTION_FLOW_ENTRIES);
  const fail = (message: string): never => {
    throw new TextDirectionPairingError({ message });
  };
  if (paired.size !== TEXT_DIRECTION_FLOW_ENTRIES.length) {
    fail("The ST_TextDirection pairing names a token twice.");
  }
  for (const token of tokens) {
    if (!paired.has(token)) {
      fail(`The ST_TextDirection pairing does not say which flow \`${token}\` names.`);
    }
  }
  for (const token of paired.keys()) {
    if (!tokens.includes(token)) {
      fail(`The ST_TextDirection pairing carries \`${token}\`, which the enumeration does not.`);
    }
  }
  const flows = tokens.filter((token) => paired.get(token) === token);
  for (const [token, flow] of paired) {
    if (!flows.includes(flow)) {
      fail(`\`${token}\` names flow \`${flow}\`, which does not name itself.`);
    }
  }
  return flows;
};

const main = async (): Promise<void> => {
  const index = buildIndex(await loadSchemaGraph());

  const paragraphAlignments = enumerationOf(index, WML_NAMESPACE, "ST_Jc");
  const tabStopAlignments = enumerationOf(index, WML_NAMESPACE, "ST_TabJc");
  const tableAlignments = enumerationOf(index, WML_NAMESPACE, "ST_JcTable");
  const textDirections = enumerationOf(index, WML_NAMESPACE, "ST_TextDirection");
  const textDirectionFlowNames = textDirectionFlows(textDirections);
  const numberFormats = enumerationOf(index, WML_NAMESPACE, "ST_NumberFormat");

  const rendered = renderModule({
    summary:
      "The model enumerations that mirror an OOXML simple type, member for\nmember, so a token the format declares cannot go unspelled.",
    script: SCRIPT,
    lists: [
      renderList({
        name: "PARAGRAPH_ALIGNMENTS",
        type: "ParagraphAlignment",
        doc: `/**
 * \`ST_Jc\`: every token a \`w:jc/@w:val\` may carry.
 *
 * \`start\` and \`end\` are the direction-aware alignments, distinct from
 * \`left\` and \`right\`: in a right-to-left paragraph \`start\` sits at the
 * right margin. They are members in their own right, not spellings of
 * \`left\` and \`right\`, and the layout resolves them against the
 * paragraph's direction. \`numTab\` aligns to the list number's tab stop.
 */`,
        members: paragraphAlignments,
      }),
      renderList({
        name: "TAB_STOP_ALIGNMENTS",
        type: "TabStopAlignment",
        doc: `/**
 * \`ST_TabJc\`: every token a \`w:tab/@w:val\` may carry.
 *
 * \`start\` and \`end\` are the direction-aware members, the same distinction
 * \`ST_Jc\` draws; \`clear\` removes an inherited stop rather than declaring
 * one, and \`num\` is the stop a numbered paragraph's text hangs from.
 */`,
        members: tabStopAlignments,
      }),
      renderList({
        name: "TABLE_ALIGNMENTS",
        type: "TableAlignment",
        doc: `/**
 * \`ST_JcTable\`: every token a table's or a row's \`w:jc/@w:val\` may carry.
 *
 * A narrower vocabulary than \`ST_Jc\`: a table is placed, not justified, so
 * there is no \`both\`. \`start\` and \`end\` are the direction-aware members,
 * the same distinction \`ST_Jc\` draws, resolved against the table's own
 * \`w:bidiVisual\` rather than a paragraph's direction.
 */`,
        members: tableAlignments,
      }),
      renderList({
        name: "TEXT_DIRECTIONS",
        type: "TextDirection",
        doc: `/**
 * \`ST_TextDirection\`: every token a \`w:textDirection/@w:val\` may carry, on
 * a table cell, a section or a paragraph.
 *
 * Twelve tokens for six flows: each flow has a short spelling and a long one
 * naming the character and line progressions in full.
 * {@link TEXT_DIRECTION_FLOWS} is the six, and
 * {@link TEXT_DIRECTION_FLOW_BY_TOKEN} pairs each token with its flow.
 */`,
        members: textDirections,
      }),
      renderList({
        name: "TEXT_DIRECTION_FLOWS",
        type: "TextDirectionFlow",
        doc: `/**
 * The six text flows \`ST_TextDirection\` names, spelled the Strict way.
 *
 * Strict's \`ST_TextDirection\` (ECMA-376 Part 1 §17.18.93) enumerates exactly
 * these six; Transitional adds a second spelling of each (Part 4 §14.11.7).
 * Rendering is decided per flow, so a cell written \`tbRl\` and its Strict twin
 * written \`rl\` paint the same.
 */`,
        members: textDirectionFlowNames,
      }),
      renderMap({
        name: "TEXT_DIRECTION_FLOW_BY_TOKEN",
        keyType: "TextDirection",
        valueType: "TextDirectionFlow",
        doc: `/**
 * The flow each \`ST_TextDirection\` token names.
 *
 * ECMA-376 Part 4 §14.11.7 gives each Transitional-only token as semantically
 * equivalent to a Strict one: \`btLr\` to \`lr\`, \`lrTb\` to \`tb\`, \`lrTbV\`
 * to \`tbV\`, \`tbLrV\` to \`lrV\`, \`tbRl\` to \`rl\` and \`tbRlV\` to \`rlV\`.
 * A reader keeps the token as authored and a writer writes it back; only
 * rendering goes through the flow.
 */`,
        entries: TEXT_DIRECTION_FLOW_ENTRIES,
      }),
      renderList({
        name: "NUMBER_FORMATS",
        type: "NumberFormat",
        doc: `/**
 * \`ST_NumberFormat\`: every token a \`w:numFmt/@w:val\` may carry, on a
 * numbering level, a note's properties or a section's page numbers.
 *
 * \`custom\` counts by the token list in the sibling \`@w:format\` rather
 * than by a vocabulary of its own; \`none\` prints no counter at all.
 */`,
        members: numberFormats,
      }),
    ],
  });

  await emitGeneratedModule({
    mode: process.argv.at(2) ?? "write",
    outputPath: OUTPUT_PATH,
    rendered,
    summary:
      `${String(paragraphAlignments.length)} paragraph alignments, ` +
      `${String(tabStopAlignments.length)} tab stop alignments, ` +
      `${String(tableAlignments.length)} table alignments, ` +
      `${String(textDirections.length)} text directions in ` +
      `${String(textDirectionFlowNames.length)} flows, ` +
      `${String(numberFormats.length)} number formats`,
    script: SCRIPT,
  });
};

await main();
