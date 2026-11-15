/**
 * A typed control's value in a canonical session: the dropdown, combo box,
 * date and checkbox widgets set the control's properties (`w:lastValue`,
 * `w:fullDate`, `w14:checked`) and replace its shown content in one journal
 * entry, so one undo takes both back and suggesting mode tracks the content.
 * The headless content-control rules decide the new value: a lock, a type
 * mismatch or a data binding refuses, and the refusal names its ground.
 */

import { panic, Result } from "better-result";
import type { EditorState } from "prosemirror-state";

import {
  blockControlHolding,
  DOCUMENT_OP_TYPES,
  type EditorIntent,
  type SetContentControlOp,
  storyBody,
} from "@stll/docx-core/ops";

import {
  type CanonicalCommit,
  type CanonicalSession,
  CanonicalSessionError,
} from "../canonical/canonicalSession";
import {
  setContentControlValue,
  type SetContentControlValueInput,
} from "../content-controls";
import {
  ContentControlBoundError,
  ContentControlLockedError,
  ContentControlTypeError,
} from "../content-controls/errors";
import {
  CANONICAL_GAP,
  CANONICAL_RULE,
  type CanonicalRefusalGround,
} from "../types/canonicalCapabilities";
import type { BlockContent } from "../types/document";
import { readModel } from "../types/documentModel";

const refuse = (
  message: string,
  ground: CanonicalRefusalGround = { gap: CANONICAL_GAP.commands },
) =>
  Result.err(new CanonicalSessionError({ ground, message, reason: "refused" }));

export const prepareContentControlValue = (
  session: CanonicalSession,
  state: EditorState,
  command: { at: number; input: SetContentControlValueInput },
): Result<CanonicalCommit, CanonicalSessionError> => {
  const node = session.projection.doc.nodeAt(command.at);
  if (node?.type.name !== "blockSdt")
    return refuse("The content control is no longer there.");
  let paragraphsOnly = node.childCount > 0;
  node.forEach((child) => {
    if (child.type.name !== "paragraph") paragraphsOnly = false;
  });
  if (!paragraphsOnly)
    return refuse("A control that holds more than paragraphs has no value.");
  // From the start of its first paragraph to the end of its last.
  const from = session.projection.addressAt(command.at + 2);
  if (from.isErr()) return from;
  const to = session.projection.addressAt(command.at + node.nodeSize - 2);
  if (to.isErr()) return to;
  const document = readModel(session.document);
  const held = blockControlHolding(
    storyBody(document, from.value.story).content,
    from.value.blockId,
  );
  if (held === undefined)
    return refuse("The content control is no longer there.");
  const { control } = held;
  const controlId = control.properties.id;
  if (controlId === undefined)
    return refuse("A content control without an id cannot be set.");
  let next: BlockContent | undefined;
  try {
    next = setContentControlValue(
      { package: { document: { content: [control] } } },
      { id: controlId },
      command.input,
    ).package.document.content.at(0);
  } catch (error) {
    if (error instanceof ContentControlLockedError)
      return refuse(error.message, {
        rule: CANONICAL_RULE.contentControlLock,
      });
    if (
      error instanceof ContentControlTypeError ||
      error instanceof ContentControlBoundError
    )
      return refuse(error.message);
    throw error;
  }
  if (next?.type !== "blockSdt")
    return panic("A control's value keeps it a block-level control.");
  const [paragraph] = next.content;
  if (paragraph?.type !== "paragraph")
    return panic("A control's value is shown in one paragraph.");
  const op: SetContentControlOp = {
    type: DOCUMENT_OP_TYPES.SET_CONTENT_CONTROL,
    story: from.value.story,
    controlId,
    blockIds: held.paragraphs.map(
      ({ paraId }) =>
        paraId ?? panic("A projected paragraph needs an identity."),
    ),
    expected: structuredClone(control.properties),
    properties: next.properties,
  };
  // The shown content goes in as pasted content would, inside the control's
  // paragraph: suggesting mode tracks it, and a checkbox keeps its symbol.
  const intent: EditorIntent = {
    type: "replaceFragment",
    from: from.value,
    to: to.value,
    paragraphs: [paragraph],
    openStart: 1,
    openEnd: 1,
  };
  return session.prepareIntents(state, {
    intents: [intent],
    resourceOps: [op],
  });
};
