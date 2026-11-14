import { panic, Result } from "better-result";
import type { EditorState } from "prosemirror-state";
import { CellSelection } from "prosemirror-tables";

import {
  headingOutlineLevel,
  NO_PARAGRAPH_NUMBERING,
  paragraphNumberingReference,
  resolveParagraphNumbering,
} from "@stll/docx-core/model";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  type EditorIntent,
  editorParagraphGroups,
  editorPositionAtPhysicalPosition,
  OP_STORIES,
  physicalPositionAtEditorOffset,
} from "@stll/docx-core/ops";

import {
  type CanonicalCommit,
  type CanonicalSession,
  CanonicalSessionError,
} from "../canonical/canonicalSession";
import {
  BUILT_IN_STYLE_NAME,
  createBuiltInStyleIndex,
} from "../docx/builtInStyles";
import { createNote } from "../docx/noteReferenceMark";
import {
  getCachedNumberingMap,
  numberingLevelUsesBulletMarker,
} from "../docx/numberingParser";
import { cloneParagraphWithoutPropertySource } from "../docx/paragraphPropertySource";
import { visitDocxParagraphs } from "../docx/paragraphTraversal";
import type { CanonicalCommandIntent } from "../prosemirror/canonicalCommands";
import { listRequestsForMarker } from "../prosemirror/listAutoformatMarkers";
import {
  type ListRequest,
  unusedListInstanceFor,
} from "../prosemirror/listNumbering";
import {
  effectiveParagraphNumbering,
  paragraphNumberingSourcesFromBoundary,
} from "../prosemirror/numberingAttr";
import { createStyleResolver } from "../prosemirror/styles/styleResolver";
import { CANONICAL_GAP } from "../types/canonicalCapabilities";
import type {
  BlockContent,
  Document,
  ListLevel,
  Paragraph,
} from "../types/document";
import { readModel } from "../types/documentModel";
import { collectHeadings } from "../utils/headingCollector";

const DEFAULT_TOC_PAGE_WIDTH_TWIPS = 12240;
const DEFAULT_TOC_MARGIN_TWIPS = 1440;

/** The section endpoint a paragraph holds now, distinguishing an explicit `undefined` from an absent key. */
const currentSectionEndpoint = (paragraph: Paragraph) => {
  const current = paragraph.sectionProperties;
  if (current !== undefined)
    return { type: "present", value: current } as const;
  if (Object.hasOwn(paragraph, "sectionProperties"))
    return { type: "undefined" } as const;
  return { type: "omitted" } as const;
};

export const CELL_SELECTION_PLACEMENT_REFUSAL =
  "A note or table of contents goes at a caret, not over selected cells.";

const refuse = (message: string) =>
  Result.err(
    new CanonicalSessionError({
      ground: { gap: CANONICAL_GAP.commands },
      message,
      reason: "refused",
    }),
  );

const noChange = (message: string) =>
  Result.err(
    new CanonicalSessionError({
      ground: { gap: CANONICAL_GAP.commands },
      message,
      reason: "noChange",
    }),
  );

const selectedParagraphs = (
  session: CanonicalSession,
  state: EditorState,
  topLevelOnly = false,
) => {
  const ids = new Set<string>();
  const paragraphs: Paragraph[] = [];
  state.doc.nodesBetween(
    state.selection.from,
    state.selection.to,
    (node, position) => {
      if (
        node.type.name !== "paragraph" ||
        typeof node.attrs["paraId"] !== "string"
      )
        return;
      if (topLevelOnly && state.doc.resolve(position).depth !== 0) return;
      const id = node.attrs["paraId"];
      if (ids.has(id)) return;
      ids.add(id);
      const paragraph = session.projection.paragraph(id)?.source;
      if (paragraph !== undefined) paragraphs.push(paragraph);
    },
  );
  return paragraphs;
};

const effectiveNumbering = (paragraph: Paragraph | undefined) =>
  effectiveParagraphNumbering(
    paragraphNumberingSourcesFromBoundary({
      numPr: paragraph?.formatting?.numPr,
      numPrFromStyle: paragraph?.formatting?.numPrFromStyle,
    }),
  );

/**
 * The numbering that takes a paragraph out of its list. A paragraph whose
 * style chain numbers it states the cancellation (`w:numId="0"`), since only
 * that overrides the style; any other paragraph loses its own numbering, so it
 * keeps no override it never had.
 */
const numberingRemoval = (session: CanonicalSession) => {
  const styles = createStyleResolver(
    readModel(session.document).package.styles,
  );
  return (paragraph: Paragraph): typeof NO_PARAGRAPH_NUMBERING | null =>
    resolveParagraphNumbering(
      styles.resolveParagraphStyle(paragraph.formatting?.styleId)
        .paragraphFormatting?.numPr,
    ).kind === "reference"
      ? NO_PARAGRAPH_NUMBERING
      : null;
};

/** Whether a paragraph is in a list, by numbering of its own or its style's. */
export const isCanonicalListParagraph = (
  paragraph: Paragraph | undefined,
): boolean => effectiveNumbering(paragraph).kind === "reference";

type ParagraphPosition = {
  container: readonly BlockContent[];
  index: number;
  paragraph: Paragraph;
};

/** Each selected paragraph with the block container that owns it. A list may run through several containers, such as table cells, as the editor applies it. */
const paragraphPositions = (
  session: CanonicalSession,
  paragraphs: readonly Paragraph[],
): Result<readonly ParagraphPosition[], CanonicalSessionError> => {
  if (paragraphs.length === 0)
    return noChange("List changes require a paragraph selection.");
  const positions: ParagraphPosition[] = [];
  for (const paragraph of paragraphs) {
    const address = session.projection.paragraph(paragraph.paraId ?? "");
    if (address === undefined)
      return refuse("A selected paragraph has no canonical address.");
    const index = address.container.findIndex((item) => item === paragraph);
    if (index < 0)
      return refuse(
        "A selected paragraph is missing from its block container.",
      );
    positions.push({ container: address.container, index, paragraph });
  }
  return Result.ok(positions);
};

/** Content a paragraph keeps no trace of once its pending changes are accepted. */
const GONE_ON_ACCEPT: ReadonlySet<string> = new Set([
  "deletion",
  "moveFrom",
  "bookmarkStart",
  "bookmarkEnd",
  "commentRangeStart",
  "commentRangeEnd",
  "moveFromRangeStart",
  "moveFromRangeEnd",
  "moveToRangeStart",
  "moveToRangeEnd",
]);

/** A paragraph accepting removes: its mark and every character are going. */
const goesOnAccept = (paragraph: Paragraph): boolean =>
  (paragraph.pPrMark?.kind === "del" ||
    paragraph.pPrMark?.kind === "moveFrom") &&
  paragraph.content.every(({ type }) => GONE_ON_ACCEPT.has(type));

/**
 * Every list instance a paragraph of the main story refers to once its pending
 * changes are accepted: directly, through its style, or in a pending property
 * change. A paragraph on its way out refers to nothing, so typing a list
 * marker over a struck list takes the instance editing would.
 */
const referencedNumIds = (document: Document): Set<number> => {
  const referenced = new Set<number>();
  visitDocxParagraphs(
    { documentBody: document.package.document },
    (paragraph) => {
      if (goesOnAccept(paragraph)) return;
      for (const numPr of [
        paragraph.formatting?.numPr,
        paragraph.formatting?.numPrFromStyle,
      ])
        if (numPr?.kind === "reference") referenced.add(numPr.numId);
      for (const change of paragraph.propertyChanges ?? [])
        if (change.previousFormatting?.numPr?.kind === "reference")
          referenced.add(change.previousFormatting.numPr.numId);
    },
  );
  return referenced;
};

const numberingIntent = (
  session: CanonicalSession,
  positions: readonly ParagraphPosition[],
  request: ListRequest,
): EditorIntent => {
  const definitions = readModel(session.document).package.numbering;
  const map =
    definitions === undefined ? undefined : getCachedNumberingMap(definitions);
  const container = positions.at(0)?.container ?? [];
  const first = positions.at(0)?.index ?? 0;
  const last = positions.at(-1)?.index ?? first;
  const neighbors = [
    first > 0 ? container.at(first - 1) : undefined,
    container.at(last + 1),
  ];
  const neighbor = neighbors.find((item) => {
    // A paragraph on its way out is no neighbour once the change is accepted.
    if (
      item?.type !== "paragraph" ||
      goesOnAccept(item) ||
      item.formatting?.numPr?.kind !== "reference" ||
      item.formatting.numPrFromStyle !== undefined
    )
      return false;
    const level = map?.getLevel(
      item.formatting.numPr.numId,
      item.formatting.numPr.ilvl ?? 0,
    );
    if (
      level == null ||
      numberingLevelUsesBulletMarker(level) !== (request.kind === "bullet")
    )
      return false;
    return (
      request.format === undefined ||
      (level.numFmt === request.format.numFmt &&
        level.lvlText === request.format.lvlText)
    );
  });
  let numId: number;
  let ilvl = 0;
  let target: Extract<EditorIntent, { type: "setList" }>["target"];
  const unused =
    map === undefined
      ? undefined
      : unusedListInstanceFor(
          map,
          referencedNumIds(readModel(session.document)),
          request,
        );
  if (
    neighbor?.type === "paragraph" &&
    neighbor.formatting?.numPr?.kind === "reference"
  ) {
    numId = neighbor.formatting.numPr.numId;
    ilvl = neighbor.formatting.numPr.ilvl ?? 0;
    target = { type: "existing", numId };
  } else if (unused !== undefined) {
    numId = unused;
    target = { type: "existing", numId };
  } else {
    numId =
      Math.max(0, ...(definitions?.nums.map((num) => num.numId) ?? [])) + 1;
    const abstractNumId =
      Math.max(
        -1,
        ...(definitions?.abstractNums.map((num) => num.abstractNumId) ?? []),
      ) + 1;
    const levels = Array.from({ length: 9 }, (_, level): ListLevel => ({
      ilvl: level,
      start: level === 0 ? (request.start ?? 1) : 1,
      numFmt:
        request.kind === "bullet"
          ? "bullet"
          : (request.format?.numFmt ?? "decimal"),
      lvlText:
        request.kind === "bullet"
          ? "•"
          : (request.format?.lvlText ?? "%1.").replaceAll(
              "%1",
              `%${level + 1}`,
            ),
      suffix: "tab",
      pPr: {
        indentLeft: (level + 1) * 720,
        indentFirstLine: 360,
        hangingIndent: true,
      },
    }));
    target = {
      type: "new",
      num: { numId, abstractNumId },
      abstractNum: { abstractNumId, multiLevelType: "multilevel", levels },
    };
  }
  return {
    type: "setList",
    target,
    items: positions.map(({ paragraph }) => ({
      at: {
        story: OP_STORIES.MAIN,
        blockId: paragraph.paraId ?? "",
        offset: 0,
      },
      ilvl,
    })),
  };
};

/** Toolbar meanings enter the same operation journal as native input. */
export const prepareCanonicalCommands = (
  session: CanonicalSession,
  state: EditorState,
  commands: readonly CanonicalCommandIntent[],
): Result<CanonicalCommit, CanonicalSessionError> => {
  // A note reference or a table of contents goes at one caret; selected
  // cells have none.
  if (
    state.selection instanceof CellSelection &&
    commands.some(
      ({ type }) =>
        type === "generateTOC" ||
        type === "insertNote" ||
        type === "deleteNote",
    )
  )
    return refuse(CELL_SELECTION_PLACEMENT_REFUSAL);
  const intents: EditorIntent[] = [];
  const sectionCommands = commands.filter(
    (command) =>
      command.type === "setSectionBreak" ||
      command.type === "removeSectionBreak",
  );
  if (sectionCommands.length > 0) {
    if (
      sectionCommands.length !== commands.length ||
      sectionCommands.length !== 1
    )
      return refuse("A section change must be prepared on its own.");
    const paragraphs = selectedParagraphs(session, state, true);
    if (paragraphs.length === 0)
      return refuse("Section breaks require a top-level document paragraph.");
    const command =
      sectionCommands.at(0) ?? panic("A section command was counted.");
    return session.prepareOperations(
      state,
      paragraphs.map((paragraph) => {
        const current = paragraph.sectionProperties;
        const expected = currentSectionEndpoint(paragraph);
        return {
          type: DOCUMENT_OP_TYPES.SET_SECTION_ENDPOINT,
          endpoint: {
            type: "paragraph" as const,
            blockId:
              paragraph.paraId ??
              panic("A projected paragraph needs an identity."),
          },
          expected,
          properties:
            command.type === "removeSectionBreak"
              ? ({ type: "omitted" } as const)
              : ({
                  type: "present",
                  value: { ...current, sectionStart: command.breakType },
                } as const),
        };
      }),
    );
  }
  const paragraphs = selectedParagraphs(session, state);
  const noteCommands = commands.filter(
    (command) => command.type === "insertNote" || command.type === "deleteNote",
  );
  if (noteCommands.length > 0) {
    if (noteCommands.length !== commands.length)
      return refuse("Note changes must be prepared on their own.");
    const ops: DocumentOp[] = [];
    const ordered = noteCommands.toReversed();
    // One allocation for the whole batch keeps the inserted notes' identities disjoint.
    const blockIdGroups = session.freshBlockIds(
      ordered.map((command) =>
        command.type === "insertNote" ? command.content.length : 0,
      ),
    );
    if (blockIdGroups.isErr()) return blockIdGroups;
    for (const [commandIndex, command] of ordered.entries()) {
      if (command.type === "deleteNote") {
        const at = session.projection.addressAt(command.from);
        if (at.isErr()) return at;
        ops.push({
          type: DOCUMENT_OP_TYPES.REMOVE_NOTE,
          at: at.value,
          story: { kind: command.kind, id: command.id },
        } as const);
        continue;
      }
      const at = session.projection.addressAt(command.at);
      if (at.isErr()) return at;
      if (command.content.some((block) => block.type !== "paragraph"))
        return refuse("Canonical note insertion requires paragraph content.");
      const blockIds =
        blockIdGroups.value[commandIndex] ??
        panic("A fresh identity group was reserved per note.");
      // Inserted note content is new text, so every paragraph takes a fresh identity.
      const content = command.content.map((block, index) => {
        if (block.type !== "paragraph")
          panic("Canonical note content was already validated.");
        return cloneParagraphWithoutPropertySource(block, {
          paraId:
            blockIds[index] ??
            panic("A fresh identity was reserved per paragraph."),
        });
      });
      ops.push({
        type: DOCUMENT_OP_TYPES.ADD_NOTE,
        at: at.value,
        note:
          command.kind === "footnote"
            ? createNote("footnote", command.id, content)
            : createNote("endnote", command.id, content),
      });
    }
    return session.prepareOperations(state, ops);
  }
  for (const command of commands) {
    switch (command.type) {
      case "generateTOC": {
        const at = session.projection.addressAt(command.at);
        if (at.isErr()) return at;
        const definitions = readModel(session.document).package.styles;
        const styles = createBuiltInStyleIndex(
          definitions?.styles ?? [],
          definitions?.docDefaults,
        );
        const headings = [];
        for (const heading of collectHeadings(session.projection.doc, styles)) {
          const address = session.projection.addressAt(heading.pmPos + 1);
          if (address.isErr()) return address;
          const outline = headingOutlineLevel(heading.level);
          if (outline?.kind !== "heading")
            return refuse("A TOC heading has an invalid outline level.");
          const styleId = styles.styleIdForTableOfContentsLevel(
            outline.level + 1,
          );
          headings.push({
            blockId: address.value.blockId,
            text: heading.text,
            level: outline.level,
            ...(styleId === undefined ? {} : { styleId }),
          });
        }
        if (headings.length === 0) break;
        const titleStyleId = styles.styleIdForBuiltInName(
          BUILT_IN_STYLE_NAME.tocHeading,
        );
        const section = readModel(
          session.document,
        ).package.document.content.find(
          (item) =>
            item.type === "paragraph" &&
            (item.sectionProperties?.pageWidth ?? 0) > 0,
        );
        const properties =
          section?.type === "paragraph" ? section.sectionProperties : undefined;
        intents.push({
          type: "generateTOC",
          at: at.value,
          title: command.title,
          headings,
          tabPosition:
            (properties?.pageWidth ?? DEFAULT_TOC_PAGE_WIDTH_TWIPS) -
            (properties?.marginLeft ?? DEFAULT_TOC_MARGIN_TWIPS) -
            (properties?.marginRight ?? DEFAULT_TOC_MARGIN_TWIPS),
          ...(titleStyleId === undefined ? {} : { titleStyleId }),
        });
        break;
      }
      case "setHyperlink":
      case "removeHyperlink":
      case "insertHyperlink": {
        const from = session.projection.addressAt(command.from);
        if (from.isErr()) return from;
        const to = session.projection.addressAt(command.to);
        if (to.isErr()) return to;
        intents.push({ ...command, from: from.value, to: to.value });
        break;
      }
      case "formatTable": {
        for (const target of command.targets) {
          const at = session.projection.addressAt(target.at);
          if (at.isErr()) return at;
          const address = { story: at.value.story, blockId: at.value.blockId };
          switch (target.scope) {
            case "table":
              intents.push({
                type: "table",
                operation: {
                  type: DOCUMENT_OP_TYPES.SET_TABLE_PROPS,
                  ...address,
                  patch: target.patch,
                },
              });
              break;
            case "row":
              intents.push({
                type: "table",
                operation: {
                  type: DOCUMENT_OP_TYPES.SET_ROW_PROPS,
                  ...address,
                  patch: target.patch,
                },
              });
              break;
            case "cell":
              intents.push({
                type: "table",
                operation: {
                  type: DOCUMENT_OP_TYPES.SET_CELL_PROPS,
                  ...address,
                  patch: target.patch,
                },
              });
              break;
            default: {
              const unreachable: never = target;
              return unreachable;
            }
          }
        }
        break;
      }
      case "editTableBlock": {
        const at = session.projection.addressAt(command.at);
        if (at.isErr()) return at;
        intents.push({
          type: "tableBlock",
          operation: {
            ...command.edit,
            story: at.value.story,
            blockId: at.value.blockId,
          },
        });
        break;
      }
      case "editTableRow": {
        const at = session.projection.addressAt(command.at);
        if (at.isErr()) return at;
        intents.push({
          type: "tableRow",
          operation: {
            ...command.edit,
            story: at.value.story,
            blockId: at.value.blockId,
          },
        });
        break;
      }
      case "editTable": {
        const at = session.projection.addressAt(command.at);
        if (at.isErr()) return at;
        intents.push({
          type: "table",
          operation: {
            ...command.edit,
            story: at.value.story,
            blockId: at.value.blockId,
          },
        });
        break;
      }
      case "insertBreak": {
        const from = session.projection.addressAt(command.from);
        if (from.isErr()) return from;
        const to = session.projection.addressAt(command.to);
        if (to.isErr()) return to;
        intents.push({
          type: "insertAtom",
          from: from.value,
          to: to.value,
          atom: { type: "break", breakType: command.breakType },
        });
        break;
      }
      case "setSectionBreak":
      case "removeSectionBreak":
        return panic("Section commands are prepared before editor intents.");
      case "insertNote":
      case "deleteNote":
        return panic("Note commands are prepared before editor intents.");
      case "formatRun": {
        const from = session.projection.addressAt(command.from);
        if (from.isErr()) return from;
        const to = session.projection.addressAt(command.to);
        if (to.isErr()) return to;
        intents.push({
          type: "formatRun",
          from: from.value,
          to: to.value,
          patch: command.patch,
        });
        break;
      }
      case "formatParagraph": {
        const at = session.projection.addressAt(command.at);
        if (at.isErr()) return at;
        intents.push({
          type: "formatParagraph",
          at: at.value,
          patch: command.patch,
        });
        break;
      }
      case "toggleList": {
        const positions = paragraphPositions(session, paragraphs);
        if (positions.isErr()) return positions;
        const definitions = readModel(session.document).package.numbering;
        const map =
          definitions === undefined
            ? undefined
            : getCachedNumberingMap(definitions);
        const already =
          paragraphs.length > 0 &&
          paragraphs.every((paragraph) => {
            const numPr = effectiveNumbering(paragraph);
            return (
              numPr?.kind === "reference" &&
              numberingLevelUsesBulletMarker(
                map?.getLevel(numPr.numId, numPr.ilvl ?? 0),
              ) ===
                (command.kind === "bullet")
            );
          });
        if (!already) {
          intents.push(
            numberingIntent(session, positions.value, {
              kind: command.kind === "bullet" ? "bullet" : "numbered",
            }),
          );
          break;
        }
        const removed = numberingRemoval(session);
        for (const paragraph of paragraphs)
          intents.push({
            type: "formatParagraph",
            at: {
              story: OP_STORIES.MAIN,
              blockId: paragraph.paraId ?? "",
              offset: 0,
            },
            patch: { numPr: removed(paragraph) },
          });
        break;
      }
      case "removeList": {
        const positions = paragraphPositions(session, paragraphs);
        if (positions.isErr()) return positions;
        const removed = numberingRemoval(session);
        for (const paragraph of paragraphs)
          intents.push({
            type: "formatParagraph",
            at: {
              story: OP_STORIES.MAIN,
              blockId: paragraph.paraId ?? "",
              offset: 0,
            },
            patch: { numPr: removed(paragraph) },
          });
        break;
      }
      case "changeListLevel": {
        const positions = paragraphPositions(session, paragraphs);
        if (positions.isErr()) return positions;
        const removed = numberingRemoval(session);
        for (const paragraph of paragraphs) {
          const numPr = effectiveNumbering(paragraph);
          if (numPr?.kind !== "reference") continue;
          const at = {
            story: OP_STORIES.MAIN,
            blockId: paragraph.paraId ?? "",
            offset: 0,
          } as const;
          const level =
            (numPr.ilvl ?? 0) + (command.direction === "increase" ? 1 : -1);
          if (level > 8)
            return noChange("The list is already at its deepest level.");
          intents.push({
            type: "formatParagraph",
            at,
            patch: {
              numPr: (() => {
                if (level < 0) return removed(paragraph);
                if (
                  paragraph.formatting?.numPr?.kind !== "reference" &&
                  paragraph.formatting?.numPrFromStyle?.kind === "reference"
                )
                  return { kind: "levelOnly", ilvl: level } as const;
                return paragraphNumberingReference({
                  numId: numPr.numId,
                  ilvl: level,
                });
              })(),
            },
          });
        }
        break;
      }
      case "restartNumbering":
      case "continueNumbering": {
        const positions = paragraphPositions(session, paragraphs);
        if (positions.isErr()) return positions;
        const source = paragraphs.at(0);
        const numPr = effectiveNumbering(source);
        const definitions = readModel(session.document).package.numbering;
        if (
          source === undefined ||
          numPr?.kind !== "reference" ||
          definitions === undefined
        )
          return noChange("Numbering changes require a list paragraph.");
        const { container, index } =
          positions.value.at(0) ?? panic("Missing paragraph position");
        const affected = container
          .slice(index)
          .filter((item): item is Paragraph => {
            if (item.type !== "paragraph") return false;
            const effective = effectiveNumbering(item);
            return (
              effective?.kind === "reference" && effective.numId === numPr.numId
            );
          });
        let targetId: number;
        let target: Extract<EditorIntent, { type: "setList" }>["target"];
        if (command.type === "continueNumbering") {
          const map = getCachedNumberingMap(definitions);
          const bullet = numberingLevelUsesBulletMarker(
            map.getLevel(numPr.numId, numPr.ilvl ?? 0),
          );
          const earlier = container.slice(0, index).findLast((item) => {
            if (item.type !== "paragraph") return false;
            const effective = effectiveNumbering(item);
            return (
              effective?.kind === "reference" &&
              effective.numId !== numPr.numId &&
              numberingLevelUsesBulletMarker(
                map.getLevel(effective.numId, effective.ilvl ?? 0),
              ) === bullet
            );
          });
          const earlierNumbering =
            earlier?.type === "paragraph"
              ? effectiveNumbering(earlier)
              : undefined;
          if (earlierNumbering?.kind !== "reference")
            return noChange("There is no preceding compatible list.");
          targetId = earlierNumbering.numId;
          target = { type: "existing", numId: targetId };
        } else {
          const start = command.start ?? 1;
          if (!Number.isSafeInteger(start) || start < 0)
            return refuse(
              "The numbering start must be a non-negative integer.",
            );
          const original = definitions.nums.find(
            (num) => num.numId === numPr.numId,
          );
          if (original === undefined)
            return refuse("The numbering instance does not exist.");
          targetId =
            Math.max(0, ...definitions.nums.map((num) => num.numId)) + 1;
          const level = numPr.ilvl ?? 0;
          target = {
            type: "new",
            num: {
              ...original,
              numId: targetId,
              levelOverrides: [
                ...(original.levelOverrides ?? []).filter(
                  ({ ilvl }) => ilvl !== level,
                ),
                {
                  ...original.levelOverrides?.find(
                    ({ ilvl }) => ilvl === level,
                  ),
                  ilvl: level,
                  startOverride: start,
                },
              ],
            },
          };
        }
        intents.push({
          type: "setList",
          target,
          items: affected.map((paragraph) => {
            const effective = effectiveNumbering(paragraph);
            return {
              at: {
                story: OP_STORIES.MAIN,
                blockId: paragraph.paraId ?? "",
                offset: 0,
              },
              ilvl: effective?.kind === "reference" ? (effective.ilvl ?? 0) : 0,
            };
          }),
        });
        break;
      }
      default: {
        const unreachable: never = command;
        return unreachable;
      }
    }
  }
  return session.prepareIntents(state, { intents });
};

type CanonicalAutoformatTrigger = { from: number; to: number; text: string };

/**
 * The rule that `typed`, already committed as typed, completes at the caret of
 * `state`: the marker's removal and the paragraph change as one atomic edit.
 *
 * The typed text is its own history step before the rule, so undoing the rule
 * (or Backspace straight after it) gives back exactly what was typed, the
 * trigger space included.
 */
export const prepareCanonicalAutoformat = (
  session: CanonicalSession,
  state: EditorState,
  typed: CanonicalAutoformatTrigger,
): Result<CanonicalCommit, CanonicalSessionError> | undefined => {
  if (typed.text !== " " || typed.from !== typed.to) return undefined;
  if (!state.selection.empty) return undefined;
  const at = session.projection.addressAt(state.selection.head);
  if (at.isErr()) return undefined;
  const document = readModel(session.document);
  const logical = editorPositionAtPhysicalPosition(document, at.value);
  const address = session.projection.paragraph(logical.blockId);
  if (
    address === undefined ||
    effectiveNumbering(address.source)?.kind === "reference"
  )
    return undefined;
  const group = editorParagraphGroups(document, logical.story).find(
    ({ blockId }) => blockId === logical.blockId,
  );
  if (group === undefined) panic("Autoformat lost its logical paragraph");
  const source = address.source;
  const prefix = group.text.slice(0, logical.offset);
  const requests = listRequestsForMarker(prefix);
  const heading = /^(?<hashes>#{1,6}) $/u.exec(prefix)?.groups?.["hashes"];
  if (requests === null && heading === undefined) return undefined;
  const start = physicalPositionAtEditorOffset(document, {
    ...logical,
    offset: 0,
  });
  const request = requests?.at(0);
  const styleId = `Heading${heading?.length ?? 1}`;
  if (request === undefined && !session.hasStyle(styleId)) return undefined;
  let paragraphIntent: EditorIntent;
  if (request === undefined) {
    paragraphIntent = {
      type: "formatParagraph",
      at: start,
      patch: { styleId },
    };
  } else {
    const positions = paragraphPositions(session, [source]);
    if (positions.isErr()) return undefined;
    paragraphIntent = numberingIntent(session, positions.value, request);
  }
  return session.prepareIntents(state, {
    intents: [
      { type: "replaceText", from: start, to: at.value, text: "" },
      paragraphIntent,
    ],
  });
};
