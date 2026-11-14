import { Result } from "better-result";

import type { Document, TrackedChangeInfo } from "../model/document";
import type { DocumentEdit } from "./edits";
import { structurallyEqual } from "./equality";
import { transferParagraphPropertySnapshot } from "./paragraphPropertySnapshot";
import {
  DOCUMENT_OP_REFUSAL_REASONS,
  DocumentOpRefusal,
  type DocumentOpRefusalReason,
} from "./refusal";
import { replaceStoryBody, storyBody } from "./stories";
import { DOCUMENT_OP_TYPES, type SetRevisionInfoOp } from "./types";

const refuse = (reason: DocumentOpRefusalReason, message: string) =>
  Result.err(
    new DocumentOpRefusal({
      reason,
      message,
      opType: DOCUMENT_OP_TYPES.SET_REVISION_INFO,
    }),
  );

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const revisionInfoOf = (value: unknown): TrackedChangeInfo | undefined =>
  isRecord(value) &&
  typeof value["id"] === "number" &&
  typeof value["author"] === "string"
    ? (value as TrackedChangeInfo)
    : undefined;

/** Every paragraph id under a value. */
const paragraphIdsIn = (value: unknown, out: Set<string>): void => {
  if (Array.isArray(value)) {
    for (const item of value) paragraphIdsIn(item, out);
    return;
  }
  if (!isRecord(value)) return;
  if (value["type"] === "paragraph" && typeof value["paraId"] === "string")
    out.add(value["paraId"]);
  for (const field of Object.values(value)) paragraphIdsIn(field, out);
};

/**
 * Write the info given for each revision id onto every record carrying that
 * id, in the story's blocks and its section view. Records keep their content
 * and their place; an unchanged subtree keeps its identity.
 */
export const setRevisionInfo = (
  document: Document,
  op: SetRevisionInfoOp,
): Result<DocumentEdit, DocumentOpRefusal> => {
  const wanted = new Map<number, TrackedChangeInfo>();
  for (const info of op.infos) {
    if (wanted.has(info.id))
      return refuse(
        DOCUMENT_OP_REFUSAL_REASONS.INVALID_OPERATION,
        `Revision ${String(info.id)} is given two infos.`,
      );
    wanted.set(info.id, info);
  }
  const previous = new Map<number, TrackedChangeInfo>();
  let disagreeing: number | undefined;
  const modified = new Set<string>();
  const rebuilt = new WeakMap<object, unknown>();
  // Rebuilt values whose every change lies in paragraphs already declared
  // modified. Any other change belongs to the nearest record holding
  // paragraphs: a row's tracked deletion, a cell's or a table's property
  // change is written with the paragraphs that row, cell or table holds.
  const declared = new WeakSet<object>();
  const rebuild = (value: unknown): unknown => {
    if (typeof value !== "object" || value === null) return value;
    const known = rebuilt.get(value);
    if (known !== undefined) return known;
    let out: unknown = value;
    if (Array.isArray(value)) {
      const items = value.map(rebuild);
      if (items.some((item, index) => item !== value[index])) {
        out = items;
        if (
          items.every(
            (item, index) =>
              item === value[index] ||
              (typeof item === "object" && item !== null && declared.has(item)),
          )
        )
          declared.add(items);
      }
    } else if (isRecord(value)) {
      let changed = false;
      let undeclared = false;
      const entries: [string, unknown][] = [];
      for (const [key, field] of Object.entries(value)) {
        const info = key === "info" ? revisionInfoOf(field) : undefined;
        const replacement =
          info === undefined ? undefined : wanted.get(info.id);
        let next: unknown;
        if (info !== undefined && replacement !== undefined) {
          const before = previous.get(info.id);
          if (before !== undefined && !structurallyEqual(before, info))
            disagreeing ??= info.id;
          previous.set(info.id, info);
          next = structurallyEqual(info, replacement) ? field : replacement;
          if (next !== field) undeclared = true;
        } else {
          next = rebuild(field);
          if (
            next !== field &&
            !(typeof next === "object" && next !== null && declared.has(next))
          )
            undeclared = true;
        }
        if (next !== field) changed = true;
        entries.push([key, next]);
      }
      if (changed) {
        // A spread keeps the record's private bindings (its enumerable
        // symbol fields), and the snapshot a paragraph is bound to follows it.
        const copy: Record<string, unknown> = { ...value };
        for (const [key, field] of entries) copy[key] = field;
        transferParagraphPropertySnapshot({ source: value, target: copy });
        out = copy;
        if (
          value["type"] === "paragraph" &&
          typeof value["paraId"] === "string"
        ) {
          modified.add(value["paraId"]);
          declared.add(copy);
        } else if (!undeclared) declared.add(copy);
        else {
          const held = new Set<string>();
          paragraphIdsIn(value, held);
          if (held.size > 0) {
            for (const id of held) modified.add(id);
            declared.add(copy);
          }
        }
      }
    }
    rebuilt.set(value, out);
    return out;
  };
  const body = storyBody(document, op.story);
  const content = rebuild(body.content) as typeof body.content;
  const sections =
    body.sections === undefined
      ? undefined
      : (rebuild(body.sections) as typeof body.sections);
  const missing = [...wanted.keys()].find((id) => !previous.has(id));
  if (missing !== undefined)
    return refuse(
      DOCUMENT_OP_REFUSAL_REASONS.STALE,
      `No revision ${String(missing)} in the story.`,
    );
  // The inverse writes one info back per id, so every record of an id must
  // carry the same one.
  if (disagreeing !== undefined)
    return refuse(
      DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
      `The records of revision ${String(disagreeing)} carry different infos.`,
    );
  if (content === body.content && sections === body.sections)
    return refuse(
      DOCUMENT_OP_REFUSAL_REASONS.NO_CHANGE,
      "The revisions already carry these infos.",
    );
  return Result.ok({
    document: replaceStoryBody({
      document,
      story: op.story,
      body: {
        ...body,
        content,
        ...(sections === undefined ? {} : { sections }),
      },
    }),
    inverse: [
      {
        type: DOCUMENT_OP_TYPES.SET_REVISION_INFO,
        story: op.story,
        infos: [...previous.values()],
      },
    ],
    touched: { modified: [...modified], inserted: [], removed: [] },
  });
};
