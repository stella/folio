import { Result } from "better-result";
import type { EditorState } from "prosemirror-state";
import { OP_STORIES, type EditorIntent, type TextPosition } from "@stll/docx-core/ops";
import { paragraphNumberingReference } from "@stll/docx-core/model";
import type { Paragraph, ListLevel } from "../types/document";
import { getCachedNumberingMap } from "../docx/numberingParser";
import { listRequestsForMarker } from "../prosemirror/listAutoformatMarkers";
import type { ListRequest } from "../prosemirror/listNumbering";
import type { CanonicalCommandIntent } from "../prosemirror/canonicalCommands";
import {
  CanonicalSessionError,
  type CanonicalCommit,
  type CanonicalSession,
} from "./canonicalSession";

const refuse = (message: string) =>
  Result.err(new CanonicalSessionError({ message, reason: "refused" }));

const selectedParagraphs = (session: CanonicalSession, state: EditorState) => {
  const ids = new Set<string>();
  state.doc.nodesBetween(state.selection.from, state.selection.to, (node) => {
    if (node.type.name === "paragraph" && typeof node.attrs["paraId"] === "string")
      ids.add(node.attrs["paraId"]);
  });
  return session.document.package.document.content.filter(
    (paragraph): paragraph is Paragraph =>
      paragraph.type === "paragraph" && ids.has(paragraph.paraId ?? ""),
  );
};

const numberingIntent = (
  session: CanonicalSession,
  paragraphs: readonly Paragraph[],
  request: ListRequest,
): EditorIntent => {
  const definitions = session.document.package.numbering;
  const map = definitions === undefined ? undefined : getCachedNumberingMap(definitions);
  const body = session.document.package.document.content;
  const first = body.findIndex((item) => item === paragraphs.at(0));
  const last = body.findIndex((item) => item === paragraphs.at(-1));
  const neighbors = [first > 0 ? body.at(first - 1) : undefined, body.at(last + 1)];
  const neighbor = neighbors.find((item) => {
    if (
      item?.type !== "paragraph" ||
      item.formatting?.numPr?.kind !== "reference" ||
      item.formatting.numPrFromStyle !== undefined
    )
      return false;
    const level = map?.getLevel(item.formatting.numPr.numId, item.formatting.numPr.ilvl ?? 0);
    if (level == null || (level.numFmt === "bullet") !== (request.kind === "bullet")) return false;
    return (
      request.format === undefined ||
      (level.numFmt === request.format.numFmt && level.lvlText === request.format.lvlText)
    );
  });
  let numId: number;
  let ilvl = 0;
  let target: Extract<EditorIntent, { type: "setList" }>["target"];
  if (neighbor?.type === "paragraph" && neighbor.formatting?.numPr?.kind === "reference") {
    numId = neighbor.formatting.numPr.numId;
    ilvl = neighbor.formatting.numPr.ilvl ?? 0;
    target = { type: "existing", numId };
  } else {
    numId = Math.max(0, ...(definitions?.nums.map((num) => num.numId) ?? [])) + 1;
    const abstractNumId =
      Math.max(-1, ...(definitions?.abstractNums.map((num) => num.abstractNumId) ?? [])) + 1;
    const levels = Array.from(
      { length: 9 },
      (_, level): ListLevel => ({
        ilvl: level,
        start: request.start ?? 1,
        numFmt: request.kind === "bullet" ? "bullet" : (request.format?.numFmt ?? "decimal"),
        lvlText:
          request.kind === "bullet"
            ? "•"
            : (request.format?.lvlText ?? "%1.").replaceAll("%1", `%${level + 1}`),
        suffix: "tab",
        pPr: { indentLeft: (level + 1) * 720, indentFirstLine: 360, hangingIndent: true },
      }),
    );
    target = {
      type: "new",
      num: { numId, abstractNumId },
      abstractNum: { abstractNumId, multiLevelType: "multilevel", levels },
    };
  }
  return {
    type: "setList",
    target,
    items: paragraphs.map((paragraph) => ({
      at: { story: OP_STORIES.MAIN, blockId: paragraph.paraId ?? "", offset: 0 },
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
  const paragraphs = selectedParagraphs(session, state);
  for (const command of commands) {
    switch (command.type) {
      case "formatRun": {
        const from = session.projection.addressAt(command.from);
        if (from.isErr()) return from;
        const to = session.projection.addressAt(command.to);
        if (to.isErr()) return to;
        intents.push({ type: "formatRun", from: from.value, to: to.value, patch: command.patch });
        break;
      }
      case "formatParagraph": {
        const at = session.projection.addressAt(command.at);
        if (at.isErr()) return at;
        intents.push({ type: "formatParagraph", at: at.value, patch: command.patch });
        break;
      }
      case "toggleList": {
        const definitions = session.document.package.numbering;
        const map = definitions === undefined ? undefined : getCachedNumberingMap(definitions);
        const already =
          paragraphs.length > 0 &&
          paragraphs.every(({ formatting }) => {
            const numPr = formatting?.numPr;
            return (
              numPr?.kind === "reference" &&
              (map?.getLevel(numPr.numId, numPr.ilvl ?? 0)?.numFmt === "bullet") ===
                (command.kind === "bullet")
            );
          });
        if (!already) {
          intents.push(
            numberingIntent(session, paragraphs, {
              kind: command.kind === "bullet" ? "bullet" : "numbered",
            }),
          );
          break;
        }
        for (const paragraph of paragraphs)
          intents.push({
            type: "formatParagraph",
            at: { story: OP_STORIES.MAIN, blockId: paragraph.paraId ?? "", offset: 0 },
            patch: { numPr: { kind: "none" } },
          });
        break;
      }
      case "removeList":
        for (const paragraph of paragraphs)
          intents.push({
            type: "formatParagraph",
            at: { story: OP_STORIES.MAIN, blockId: paragraph.paraId ?? "", offset: 0 },
            patch: { numPr: { kind: "none" } },
          });
        break;
      case "changeListLevel":
        for (const paragraph of paragraphs) {
          const numPr = paragraph.formatting?.numPr;
          const at = {
            story: OP_STORIES.MAIN,
            blockId: paragraph.paraId ?? "",
            offset: 0,
          } as const;
          if (numPr?.kind !== "reference") {
            const indent = Math.max(
              0,
              (paragraph.formatting?.indentLeft ?? 0) +
                (command.direction === "increase" ? 720 : -720),
            );
            intents.push({ type: "formatParagraph", at, patch: { indentLeft: indent } });
            continue;
          }
          const level = (numPr.ilvl ?? 0) + (command.direction === "increase" ? 1 : -1);
          if (level > 8) return refuse("The list is already at its deepest level.");
          intents.push({
            type: "formatParagraph",
            at,
            patch: {
              numPr:
                level < 0
                  ? { kind: "none" }
                  : paragraphNumberingReference({ numId: numPr.numId, ilvl: level }),
            },
          });
        }
        break;
      case "restartNumbering":
      case "continueNumbering": {
        const source = paragraphs.at(0);
        const numPr = source?.formatting?.numPr;
        const definitions = session.document.package.numbering;
        if (source === undefined || numPr?.kind !== "reference" || definitions === undefined)
          return refuse("Numbering changes require a list paragraph.");
        const body = session.document.package.document.content;
        const index = body.findIndex((item) => item === source);
        const affected = body
          .slice(index)
          .filter(
            (item): item is Paragraph =>
              item.type === "paragraph" &&
              item.formatting?.numPr?.kind === "reference" &&
              item.formatting.numPr.numId === numPr.numId,
          );
        let targetId: number;
        let target: Extract<EditorIntent, { type: "setList" }>["target"];
        if (command.type === "continueNumbering") {
          const map = getCachedNumberingMap(definitions);
          const bullet = map.getLevel(numPr.numId, numPr.ilvl ?? 0)?.numFmt === "bullet";
          const earlier = body
            .slice(0, index)
            .findLast(
              (item) =>
                item.type === "paragraph" &&
                item.formatting?.numPr?.kind === "reference" &&
                item.formatting.numPr.numId !== numPr.numId &&
                (map.getLevel(item.formatting.numPr.numId, item.formatting.numPr.ilvl ?? 0)
                  ?.numFmt ===
                  "bullet") ===
                  bullet,
            );
          if (earlier?.type !== "paragraph" || earlier.formatting?.numPr?.kind !== "reference")
            return refuse("There is no preceding compatible list.");
          targetId = earlier.formatting.numPr.numId;
          target = { type: "existing", numId: targetId };
        } else {
          const start = command.start ?? 1;
          if (!Number.isSafeInteger(start) || start < 0)
            return refuse("The numbering start must be a non-negative integer.");
          const original = definitions.nums.find((num) => num.numId === numPr.numId);
          if (original === undefined) return refuse("The numbering instance does not exist.");
          targetId = Math.max(0, ...definitions.nums.map((num) => num.numId)) + 1;
          const level = numPr.ilvl ?? 0;
          target = {
            type: "new",
            num: {
              ...original,
              numId: targetId,
              levelOverrides: [
                ...(original.levelOverrides ?? []).filter(({ ilvl }) => ilvl !== level),
                {
                  ...original.levelOverrides?.find(({ ilvl }) => ilvl === level),
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
          items: affected.map((paragraph) => ({
            at: { story: OP_STORIES.MAIN, blockId: paragraph.paraId ?? "", offset: 0 },
            ilvl:
              paragraph.formatting?.numPr?.kind === "reference"
                ? (paragraph.formatting.numPr.ilvl ?? 0)
                : 0,
          })),
        });
        break;
      }
      default: {
        const unreachable: never = command;
        return unreachable;
      }
    }
  }
  return session.prepareIntents(state, intents);
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
  const source = session.projection.paragraph(at.value.blockId)?.source;
  if (source === undefined || source.formatting?.numPr?.kind === "reference") return undefined;
  const prefix =
    state.doc.resolve(from).parent.textBetween(0, at.value.offset, "", "\uFFFC") + text;
  const requests = listRequestsForMarker(prefix);
  const heading = /^(?<hashes>#{1,6}) $/u.exec(prefix)?.groups?.["hashes"];
  if (requests === null && heading === undefined) return undefined;
  const start: TextPosition = { ...at.value, offset: 0 };
  const request = requests?.at(0);
  const styleId = `Heading${heading?.length ?? 1}`;
  if (request === undefined && !session.hasStyle(styleId)) return undefined;
  return session.prepareIntents(state, [
    { type: "replaceText", from: start, to: at.value, text: "" },
    request === undefined
      ? { type: "formatParagraph", at: start, patch: { styleId } }
      : numberingIntent(session, [source], request),
  ]);
};
