/**
 * Live transaction save census. Run with:
 * bun packages/core/scripts/selective-save-edit-census.ts [--limit-per-source=10]
 *
 * Zero selects the whole cached tier-1 corpus. Fixtures contain generated
 * main-part markup; external relationship targets are not synthesized.
 * Missing files and failed parses are
 * counted explicitly. XML losses count namespace-qualified child occurrences
 * outside the edited top-level paragraph (and its split/insert descendants).
 * This is a preservation measurement, not XML schema validation.
 */
import path from "node:path";
import JSZip from "jszip";
import { TaggedError } from "better-result";
import { EditorState, TextSelection } from "prosemirror-state";
import type { Node as PMNode } from "prosemirror-model";
import { parseDocx } from "@stll/folio-core/docx/parser";
import { ensureParaIds } from "@stll/folio-core/docx/ensureParaIds";
import { createEmptyDocx, repackDocx } from "@stll/folio-core/docx/rezip";
import { attemptSelectiveSave } from "@stll/folio-core/docx/selectiveSave";
import {
  parseXmlDocument,
  getLocalName,
  getNamespaceUri,
  type XmlElement,
} from "@stll/folio-core/docx/xmlParser";
import { fromProseDoc } from "@stll/folio-core/prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "@stll/folio-core/prosemirror/conversion/toProseDoc";
import { splitBlockClearBorders } from "@stll/folio-core/prosemirror/extensions/features/BaseKeymapExtension";
import {
  ParaIdAllocatorExtension,
  ensureParaIdsInState,
} from "@stll/folio-core/prosemirror/extensions/features/ParaIdAllocatorExtension";
import {
  ParagraphChangeTrackerExtension,
  paragraphChangeTrackerKey,
} from "@stll/folio-core/prosemirror/extensions/features/ParagraphChangeTrackerExtension";
import { loadCorpusLock, sourceCheckoutPath } from "../../../scripts/lib/corpus-manifest";
import { allSubjects } from "../../../scripts/container-survival-census";
import { buildFixture } from "../../../scripts/lib/container-survival/fixture";
import {
  loadContainerSpace,
  childSlots,
  qualify,
  WML_NAMESPACE,
} from "../../../scripts/lib/container-survival/schemaSpace";
import { subjectKey } from "../../../scripts/lib/container-survival/laws";

class EditCensusError extends TaggedError("EditCensusError")<{
  message: string;
}> {}

const EDITS = [
  "type-text",
  "enter-mid-paragraph",
  "delete-paragraph",
  "paste-three-paragraphs",
] as const;
type Edit = (typeof EDITS)[number];
const argument = process.argv.find((value) => value.startsWith("--limit-per-source="));
const limit = argument === undefined ? 10 : Number(argument.split("=").at(1));
if (!Number.isSafeInteger(limit) || limit < 0)
  throw new EditCensusError({ message: "Invalid corpus sample limit" });
const space = await loadContainerSpace();
const declaredPairs = new Set(
  childSlots(space).map(
    ({ container, child }) => `${qualify(container.element)}/${qualify(child)}`,
  ),
);
const qualified = (node: XmlElement): string =>
  qualify({ namespace: getNamespaceUri(node) ?? "", name: getLocalName(node.name ?? "") });

// Scope is unchanged body content. Keep the body itself so its untouched
// container children are counted, but omit the deliberately edited paragraphs.
const inventory = (xml: string, skipped: ReadonlySet<number>): Map<string, number> => {
  const counts = new Map<string, number>();
  const visit = (node: XmlElement): void => {
    let bodyParagraph = 0;
    for (const child of node.elements ?? []) {
      if (child.type !== "element") continue;
      if (
        qualified(node) === `{${WML_NAMESPACE}}body` &&
        qualified(child) === `{${WML_NAMESPACE}}p`
      ) {
        const ordinal = bodyParagraph++;
        if (skipped.has(ordinal)) continue;
      }
      const key = `${qualified(node)}/${qualified(child)}`;
      if (declaredPairs.has(key)) counts.set(key, (counts.get(key) ?? 0) + 1);
      visit(child);
    }
  };
  const root = parseXmlDocument(xml);
  if (root !== null) visit(root);
  return counts;
};

const targets = (doc: PMNode) => {
  const found: { node: PMNode; pos: number; ordinal: number }[] = [];
  let ordinal = 0;
  doc.forEach((node, pos) => {
    if (node.type.name !== "paragraph") return;
    if (node.textContent.length >= 4) found.push({ node, pos, ordinal });
    ordinal += 1;
  });
  return found;
};

type Total = {
  measured: number;
  selective: number;
  repack: number;
  errors: number;
  lossDocuments: number;
  lostOccurrences: number;
  losses: Record<string, number>;
};
const totals = new Map<string, Total>();
const failures: { source: string; edit?: Edit; error: string }[] = [];
let missing = 0;
let ineligible = 0;
let considered = 0;
const totalFor = (group: string, edit: Edit): Total => {
  const key = `${group}/${edit}`;
  let total = totals.get(key);
  if (total === undefined) {
    total = {
      measured: 0,
      selective: 0,
      repack: 0,
      errors: 0,
      lossDocuments: 0,
      lostOccurrences: 0,
      losses: {},
    };
    totals.set(key, total);
  }
  return total;
};

type MeasureOptions = { source: string; group: string; buffer: ArrayBuffer | Uint8Array };

const measure = async ({ source, group, buffer }: MeasureOptions): Promise<void> => {
  considered += 1;
  try {
    const sourceBuffer = buffer instanceof Uint8Array ? buffer.slice().buffer : buffer;
    const parsed = await parseDocx(sourceBuffer);
    const pm = toProseDoc(parsed);
    const target = targets(pm).at(0);
    if (target === undefined || pm.childCount < 2) {
      ineligible += 1;
      return;
    }
    const zip = await JSZip.loadAsync(buffer);
    const originalXml = await zip.file("word/document.xml")?.async("text");
    if (originalXml === undefined) {
      ineligible += 1;
      return;
    }
    const before = inventory(originalXml, new Set([target.ordinal]));
    for (const edit of EDITS) {
      const total = totalFor(group, edit);
      try {
        const allocator =
          ParaIdAllocatorExtension().onSchemaReady({ schema: pm.type.schema }).plugins ?? [];
        const tracker =
          ParagraphChangeTrackerExtension().onSchemaReady({ schema: pm.type.schema }).plugins ?? [];
        const initial = ensureParaIdsInState(
          EditorState.create({ doc: pm, plugins: [...allocator, ...tracker] }),
        );
        const middle = target.pos + 1 + Math.floor(target.node.content.size / 2);
        let tr = initial.tr;
        const skipped = new Set([target.ordinal]);
        switch (edit) {
          case "type-text":
            tr.insertText("X", middle);
            break;
          case "enter-mid-paragraph": {
            const selected = initial.apply(
              initial.tr.setSelection(TextSelection.create(initial.doc, middle)),
            );
            if (
              !splitBlockClearBorders(selected, (transaction) => {
                tr = transaction;
              })
            ) {
              throw new EditCensusError({ message: "Enter command declined the target selection" });
            }
            if (tr.doc.childCount !== initial.doc.childCount + 1) {
              throw new EditCensusError({
                message: "Enter command did not create one top-level paragraph",
              });
            }
            skipped.add(target.ordinal + 1);
            break;
          }
          case "delete-paragraph":
            tr.delete(target.pos, target.pos + target.node.nodeSize);
            skipped.clear();
            break;
          case "paste-three-paragraphs":
            tr.insert(
              target.pos + target.node.nodeSize,
              [0, 1, 2].map((index) =>
                target.node.type.create(null, pm.type.schema.text(`Pasted paragraph ${index + 1}`)),
              ),
            );
            for (let index = 1; index <= 3; index++) skipped.add(target.ordinal + index);
            break;
        }
        const state = initial.apply(tr);
        const tracked = paragraphChangeTrackerKey.getState(state);
        if (tracked === undefined)
          throw new EditCensusError({ message: "Missing paragraph change tracker" });
        const document = fromProseDoc(state.doc, parsed);
        const selective = await attemptSelectiveSave(document, sourceBuffer, tracked);
        const saved = selective ?? (await repackDocx(document, { updateModifiedDate: false }));
        const savedZip = await JSZip.loadAsync(saved);
        const xml = await savedZip.file("word/document.xml")?.async("text");
        if (xml === undefined) throw new EditCensusError({ message: "Saved main part missing" });
        const after = inventory(xml, skipped);
        let losses = 0;
        for (const [pair, count] of before) {
          const loss = Math.max(0, count - (after.get(pair) ?? 0));
          if (loss === 0) continue;
          total.losses[pair] = (total.losses[pair] ?? 0) + loss;
          losses += loss;
        }
        total.measured += 1;
        total[selective === null ? "repack" : "selective"] += 1;
        total.lostOccurrences += losses;
        if (losses > 0) total.lossDocuments += 1;
      } catch (error) {
        total.errors += 1;
        failures.push({ source, edit, error: String(error) });
      }
    }
  } catch (error) {
    failures.push({ source, error: String(error) });
  }
};

// Schema-generated main-part fixtures retain the tested container beside the
// editing target, so deleting that target cannot count as losing the subject.
const base = await createEmptyDocx();
let fixtureCount = 0;
for (const subject of allSubjects(space)) {
  if (subject.kind !== "child") continue;
  if (
    !["pPr", "rPr", "tblPr", "trPr", "tcPr", "sectPr", "body"].includes(
      subject.slot.container.element.name,
    )
  )
    continue;
  const built = buildFixture(space, subject);
  if (built.status !== "built" || built.fixture.part.path !== "word/document.xml") continue;
  const zip = await JSZip.loadAsync(base);
  const seed =
    "<w:p><w:r><w:t>Editable paragraph one</w:t></w:r></w:p><w:p><w:r><w:t>Untouched paragraph two</w:t></w:r></w:p>";
  zip.file("word/document.xml", built.fixture.documentXml.replace("<w:body>", `<w:body>${seed}`));
  const buffer = await zip.generateAsync({ type: "arraybuffer" });
  await measure({ source: `fixture:${subjectKey(subject)}`, group: "fixtures-raw", buffer });
  await measure({
    source: `fixture:${subjectKey(subject)}`,
    group: "fixtures-normalized",
    buffer: (await ensureParaIds(buffer)).docx,
  });
  fixtureCount += 1;
}
const lock = await loadCorpusLock();
for (const source of lock.sources) {
  if (source.tier !== 1) continue;
  const selected = limit === 0 ? source.files : source.files.slice(0, limit);
  for (const file of selected) {
    const input = Bun.file(path.join(sourceCheckoutPath(source.id), file.path));
    if (!(await input.exists())) {
      missing += 1;
      continue;
    }
    const buffer = await input.arrayBuffer();
    await measure({ source: `${source.id}/${file.path}`, group: "corpus-raw", buffer });
    try {
      await measure({
        source: `${source.id}/${file.path}`,
        group: "corpus-normalized",
        buffer: (await ensureParaIds(buffer)).docx,
      });
    } catch (error) {
      failures.push({ source: `${source.id}/${file.path}`, error: `normalize: ${String(error)}` });
    }
  }
}
console.log(
  JSON.stringify(
    {
      limitPerSource: limit,
      fixtureCount,
      considered,
      missing,
      ineligible,
      totals: Object.fromEntries(totals),
      failures,
    },
    null,
    2,
  ),
);
