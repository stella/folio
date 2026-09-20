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
 *   ParagraphAlignment       w:ST_Jc              `w:jc/@w:val`
 *   TabStopAlignment         w:ST_TabJc           `w:tab/@w:val`
 *
 * Usage:
 *   bun scripts/generate-ooxml-enumerations.ts write
 *   bun scripts/generate-ooxml-enumerations.ts check
 */

import path from "node:path";

import { buildIndex, loadSchemaGraph, WML_NAMESPACE } from "./lib/ooxml-schema-graph";
import {
  emitGeneratedModule,
  enumerationOf,
  renderList,
  renderModule,
} from "./lib/generated-enumeration";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const OUTPUT_PATH = path.join(REPO_ROOT, "packages/docx-core/src/model/ooxmlEnumerations.gen.ts");
const SCRIPT = "generate:ooxml-enumerations";

const main = async (): Promise<void> => {
  const index = buildIndex(await loadSchemaGraph());

  const paragraphAlignments = enumerationOf(index, WML_NAMESPACE, "ST_Jc");
  const tabStopAlignments = enumerationOf(index, WML_NAMESPACE, "ST_TabJc");

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
    ],
  });

  await emitGeneratedModule({
    mode: process.argv.at(2) ?? "write",
    outputPath: OUTPUT_PATH,
    rendered,
    summary:
      `${String(paragraphAlignments.length)} paragraph alignments, ` +
      `${String(tabStopAlignments.length)} tab stop alignments`,
    script: SCRIPT,
  });
};

await main();
