import { panic } from "better-result";

import {
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  paragraphsThroughTextBoxes,
  type ReplaceInlineOp,
} from "@stll/docx-core/ops";

import { parseFieldType } from "../docx/fieldParser";
import { withInlineContentRewritten } from "../docx/paragraphTraversal";
import type {
  ComplexField,
  Document,
  Paragraph,
  ParagraphContent,
  Run,
  SimpleField,
} from "../types/document";
import type { FolioDerivedFields, FolioFieldInsertion } from "../types/fields";

type Field = SimpleField | ComplexField;

const isField = <Item extends ParagraphContent>(
  item: Item,
): item is Item & Field =>
  item.type === "simpleField" || item.type === "complexField";

/** Records whose content the final document leaves out. */
const isRemoval = (item: ParagraphContent): boolean =>
  item.type === "deletion" || item.type === "moveFrom";

const runsText = (runs: readonly ParagraphContent[]): string =>
  runs
    .map((item) =>
      item.type === "run"
        ? item.content
            .map((content) => (content.type === "text" ? content.text : ""))
            .join("")
        : "",
    )
    .join("");

/** What a field shows: its cached result. */
const displayTextOf = (field: Field): string =>
  field.type === "simpleField"
    ? runsText(field.content)
    : runsText(field.fieldResult);

/** The result run of a field showing `text`, formatted as its first result run. */
const resultRun = (field: Field, text: string): Run => {
  const first = (
    field.type === "simpleField" ? field.content : field.fieldResult
  ).find((item): item is Run => item.type === "run");
  return {
    type: "run",
    ...(first?.formatting === undefined
      ? {}
      : { formatting: first.formatting }),
    content: [{ type: "text", text }],
  };
};

const withDisplayText = <F extends Field>(field: F, text: string): F =>
  field.type === "simpleField"
    ? { ...field, content: [resultRun(field, text)] }
    : { ...field, fieldResult: [resultRun(field, text)] };

/**
 * Visit the owned fields of `content` in document order, with whether a
 * removal holds them; `replace` may return a field to put in its place.
 */
/** Returns the field to keep in place of `field`, of its own kind. */
type FieldReplacer = <F extends Field>(field: F, removed: boolean) => F;

const mapOwnedFields = <Item extends ParagraphContent>(
  content: Item[],
  owns: (instruction: string) => boolean,
  replace: FieldReplacer,
  removed = false,
): Item[] => {
  let changed = false;
  const mapped = content.map((item) => {
    if (isField(item) && owns(item.instruction)) {
      const next = replace(item, removed);
      if (next !== item) changed = true;
      return next;
    }
    const next = withInlineContentRewritten(item, (children) =>
      mapOwnedFields(children, owns, replace, removed || isRemoval(item)),
    );
    if (next !== item) changed = true;
    return next;
  });
  return changed ? mapped : content;
};

/**
 * The operations that bring the owned fields of a document's main story to
 * the display texts the host derives for them: one whole-paragraph
 * replacement per paragraph a field of which shows something else. None when
 * every field already shows its text.
 */
export const derivedFieldOps = (
  document: Document,
  { owns, displayTexts }: FolioDerivedFields,
): ReplaceInlineOp[] => {
  const paragraphs = paragraphsThroughTextBoxes(document, OP_STORIES.MAIN).map(
    ({ paragraph }) => paragraph,
  );
  const instructions: string[] = [];
  for (const paragraph of paragraphs)
    mapOwnedFields(
      paragraph.content,
      owns,
      <F extends Field>(field: F, removed: boolean): F => {
        if (!removed) instructions.push(field.instruction);
        return field;
      },
    );
  if (instructions.length === 0) return [];
  const texts = displayTexts(instructions);
  if (texts.length !== instructions.length)
    panic(
      `derivedFields.displayTexts returned ${texts.length} texts for ${instructions.length} fields.`,
    );
  // The texts in field order, taken one per field outside a removal.
  const pending = texts.values();
  const assign = <F extends Field>(field: F, removed: boolean): F => {
    if (removed) return field;
    const text = pending.next();
    if (text.done === true) return panic("A derived field text is missing.");
    return displayTextOf(field) === text.value
      ? field
      : withDisplayText(field, text.value);
  };
  const ops: ReplaceInlineOp[] = [];
  for (const paragraph of paragraphs) {
    const content = mapOwnedFields(paragraph.content, owns, assign);
    if (content === paragraph.content) continue;
    if (paragraph.paraId === undefined)
      panic("A canonical paragraph holding a field has no paraId.");
    ops.push({
      type: DOCUMENT_OP_TYPES.REPLACE_INLINE,
      story: OP_STORIES.MAIN,
      blockId: paragraph.paraId,
      expected: paragraph.content,
      content,
    });
  }
  return ops;
};

/** The paragraph fragment an inserted field goes in as: the field alone. */
export const fieldInsertionParagraph = ({
  instruction,
  displayText,
  locked,
}: FolioFieldInsertion): Paragraph => ({
  type: "paragraph",
  content: [
    {
      type: "simpleField",
      instruction,
      fieldType: parseFieldType(instruction),
      content: [
        { type: "run", content: [{ type: "text", text: displayText }] },
      ],
      ...(locked === undefined ? {} : { fldLock: locked }),
    },
  ],
});
