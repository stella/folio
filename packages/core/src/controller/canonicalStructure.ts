import { panic, Result } from "better-result";
import type { EditorState } from "prosemirror-state";

import {
  headingOutlineLevel,
  paragraphNumberingReference,
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
  BUILT_IN_STYLE_NAME,
  createBuiltInStyleIndex,
} from "../docx/builtInStyles";
import { createNote } from "../docx/noteReferenceMark";
import {
  getCachedNumberingMap,
  numberingLevelUsesBulletMarker,
} from "../docx/numberingParser";
import { cloneParagraphWithoutPropertySource } from "../docx/paragraphPropertySource";
import type { CanonicalCommandIntent } from "../prosemirror/canonicalCommands";
import { listRequestsForMarker } from "../prosemirror/listAutoformatMarkers";
import type { ListRequest } from "../prosemirror/listNumbering";
import {
  effectiveParagraphNumbering,
  paragraphNumberingSourcesFromBoundary,
} from "../prosemirror/numberingAttr";
import { CANONICAL_GAP } from "../types/canonicalCapabilities";
import type { BlockContent, ListLevel, Paragraph } from "../types/document";
import { collectHeadings } from "../utils/headingCollector";
import {
  type CanonicalCommit,
  type CanonicalSession,
  CanonicalSessionError,
} from "./canonicalSession";

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

const refuse = (message: string) =>
  Result.err(
    new CanonicalSessionError({
      gap: CANONICAL_GAP.commands,
      message,
      reason: "refused",
    }),
  );

const noChange = (message: string) =>
  Result.err(
    new CanonicalSessionError({
      gap: CANONICAL_GAP.commands,
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

type ParagraphPosition = {
  container: readonly BlockContent[];
  index: number;
  paragraph: Paragraph;
};

/** List structure is local to the block container that owns each paragraph. */
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
  const container = positions.at(0)?.container;
  if (positions.some((position) => position.container !== container))
    return refuse("A list command cannot span multiple block containers.");
  return Result.ok(positions);
};

const numberingIntent = (
  session: CanonicalSession,
  positions: readonly ParagraphPosition[],
  request: ListRequest,
): EditorIntent => {
  const definitions = session.document.package.numbering;
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
    if (
      item?.type !== "paragraph" ||
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
  if (
    neighbor?.type === "paragraph" &&
    neighbor.formatting?.numPr?.kind === "reference"
  ) {
    numId = neighbor.formatting.numPr.numId;
    ilvl = neighbor.formatting.numPr.ilvl ?? 0;
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
        const definitions = session.document.package.styles;
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
        const section = session.document.package.document.content.find(
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
        const definitions = session.document.package.numbering;
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
        for (const paragraph of paragraphs)
          intents.push({
            type: "formatParagraph",
            at: {
              story: OP_STORIES.MAIN,
              blockId: paragraph.paraId ?? "",
              offset: 0,
            },
            patch: { numPr: { kind: "none" } },
          });
        break;
      }
      case "removeList":
        {
          const positions = paragraphPositions(session, paragraphs);
          if (positions.isErr()) return positions;
        }
        for (const paragraph of paragraphs)
          intents.push({
            type: "formatParagraph",
            at: {
              story: OP_STORIES.MAIN,
              blockId: paragraph.paraId ?? "",
              offset: 0,
            },
            patch: { numPr: { kind: "none" } },
          });
        break;
      case "changeListLevel":
        {
          const positions = paragraphPositions(session, paragraphs);
          if (positions.isErr()) return positions;
        }
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
                if (level < 0) return { kind: "none" } as const;
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
      case "restartNumbering":
      case "continueNumbering": {
        const positions = paragraphPositions(session, paragraphs);
        if (positions.isErr()) return positions;
        const source = paragraphs.at(0);
        const numPr = effectiveNumbering(source);
        const definitions = session.document.package.numbering;
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

type CanonicalAutoformatInput = { from: number; to: number; text: string };

/** Rules compile their replacement and paragraph change as one atomic edit. */
export const prepareCanonicalAutoformat = (
  session: CanonicalSession,
  state: EditorState,
  { from, to, text }: CanonicalAutoformatInput,
): Result<CanonicalCommit, CanonicalSessionError> | undefined => {
  if (text !== " " || from !== to) return undefined;
  const at = session.projection.addressAt(from);
  if (at.isErr()) return undefined;
  const logical = editorPositionAtPhysicalPosition(session.document, at.value);
  const address = session.projection.paragraph(logical.blockId);
  if (
    address === undefined ||
    effectiveNumbering(address.source)?.kind === "reference"
  )
    return undefined;
  const group = editorParagraphGroups(session.document, logical.story).find(
    ({ blockId }) => blockId === logical.blockId,
  );
  if (group === undefined) panic("Autoformat lost its logical paragraph");
  const source = address.source;
  const prefix = group.text.slice(0, logical.offset) + text;
  const requests = listRequestsForMarker(prefix);
  const heading = /^(?<hashes>#{1,6}) $/u.exec(prefix)?.groups?.["hashes"];
  if (requests === null && heading === undefined) return undefined;
  const start = physicalPositionAtEditorOffset(session.document, {
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
