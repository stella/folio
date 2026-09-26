import type { Node as PMNode } from "prosemirror-model";
import {
  AddMarkStep,
  AddNodeMarkStep,
  AttrStep,
  RemoveMarkStep,
  RemoveNodeMarkStep,
  type Step,
} from "prosemirror-transform";

type RecordNodeResolutionOptions = {
  before: PMNode;
  after: PMNode;
  position: number;
  steps: Step[];
};
/** Record only carrier edits; content replacements are recorded by their owning walk. */
export const recordNodeResolution = ({
  before,
  after,
  position,
  steps,
}: RecordNodeResolutionOptions): void => {
  if (before === after) return;
  if (!before.isText) {
    for (const name of Object.keys(after.attrs)) {
      if (before.attrs[name] !== after.attrs[name])
        steps.push(new AttrStep(position, name, after.attrs[name]));
    }
  }
  for (const mark of before.marks) {
    if (mark.isInSet(after.marks)) continue;
    steps.push(
      before.isText
        ? new RemoveMarkStep(position, position + before.nodeSize, mark)
        : new RemoveNodeMarkStep(position, mark),
    );
  }
  for (const mark of after.marks) {
    if (mark.isInSet(before.marks)) continue;
    steps.push(
      before.isText
        ? new AddMarkStep(position, position + before.nodeSize, mark)
        : new AddNodeMarkStep(position, mark),
    );
  }
};
