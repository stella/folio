/**
 * Put the text of a DrawingML group's text boxes back into the group.
 *
 * The parser lifts each text box out of a group into a text box shape right
 * after the group's drawing, and the drawing replays the group's authored XML.
 * Written as they stand, the two would put every text box in the document
 * twice: once inside the group and once as a drawing of its own. So before a
 * paragraph is written, each lifted text box is matched to the group it came
 * from, its content is written into that group's XML when it changed, and it
 * writes nothing itself.
 *
 * A text box whose group is no longer in the paragraph — deleted, or the box
 * copied elsewhere — is written as the ordinary text box it now is.
 */

import type {
  DrawingContent,
  Hyperlink,
  InlineSdt,
  Paragraph,
  ParagraphContent,
  Run,
  RunContent,
  ShapeContent,
  ShapeTextBody,
  TrackedRunChange,
} from "../../types/document";
import {
  groupTextContentFingerprint,
  groupXmlFingerprint,
  replaceGroupTextBoxContent,
} from "../drawingGroupChildren";
import type { GroupTextBoxEdit } from "../drawingGroupChildren";

type SerializeTextBody = (blocks: ShapeTextBody["content"]) => string;

type GroupDrawing = { drawing: DrawingContent; fingerprint: string };

/** Everything a run can sit in on the way down from a paragraph. */
type WalkedContent =
  | ParagraphContent
  | InlineSdt["content"][number]
  | TrackedRunChange["content"][number]
  | Hyperlink["children"][number];

/** Every run in document order, through the containers a text box is lifted from. */
const forEachRun = (content: readonly WalkedContent[], onRun: (run: Run) => void): void => {
  for (const item of content) {
    switch (item.type) {
      case "run":
        onRun(item);
        break;
      case "hyperlink":
        forEachRun(item.children, onRun);
        break;
      case "inlineSdt":
      case "insertion":
      case "deletion":
      case "moveFrom":
      case "moveTo":
        forEachRun(item.content, onRun);
        break;
      default:
        break;
    }
  }
};

/** The same walk, rebuilding each container whose runs changed. */
const mapRuns = <T extends WalkedContent>(content: readonly T[], mapRun: (run: Run) => Run): T[] =>
  content.map((item): T => {
    switch (item.type) {
      case "run":
        return mapRun(item) as T;
      case "hyperlink":
        return { ...item, children: mapRuns(item.children, mapRun) } as T;
      case "inlineSdt":
      case "insertion":
      case "deletion":
      case "moveFrom":
      case "moveTo":
        return { ...item, content: mapRuns(item.content, mapRun) } as T;
      default:
        return item;
    }
  });

const isGroupDrawing = (content: RunContent): content is DrawingContent & { rawXml: string } =>
  content.type === "drawing" && content.rawXml !== undefined && content.rawXml.includes("wgp");

const isGroupChild = (content: RunContent): content is ShapeContent =>
  content.type === "shape" && content.shape.groupChild !== undefined;

/**
 * The paragraph content as it should be written: lifted text boxes dropped,
 * and the drawing of each group whose text changed carrying the new text.
 *
 * Returns the paragraph's own content when it holds no lifted text box.
 */
export const writeGroupTextBoxesBack = (
  paragraph: Paragraph,
  serializeTextBody: SerializeTextBody,
): readonly ParagraphContent[] => {
  const members: { member: ShapeContent; group: GroupDrawing | undefined }[] = [];
  const groups: GroupDrawing[] = [];
  forEachRun(paragraph.content, (run) => {
    for (const content of run.content) {
      if (isGroupDrawing(content)) {
        groups.push({ drawing: content, fingerprint: groupXmlFingerprint(content.rawXml) });
      } else if (isGroupChild(content)) {
        const wanted = content.shape.groupChild?.group;
        members.push({
          member: content,
          group: groups.findLast((candidate) => candidate.fingerprint === wanted),
        });
      }
    }
  });
  if (members.length === 0) {
    return paragraph.content;
  }

  const absorbed = new Set<RunContent>();
  const editsByDrawing = new Map<DrawingContent, GroupTextBoxEdit[]>();
  const claimedPaths = new Map<DrawingContent, Set<string>>();
  for (const { member, group } of members) {
    const groupChild = member.shape.groupChild;
    if (!group || !groupChild) {
      continue;
    }
    const pathKey = groupChild.path.join("/");
    const claimed = claimedPaths.get(group.drawing) ?? new Set<string>();
    if (claimed.has(pathKey)) {
      // A copy of a text box already written back is a text box of its own.
      continue;
    }
    claimed.add(pathKey);
    claimedPaths.set(group.drawing, claimed);
    absorbed.add(member);
    const content = member.shape.textBody?.content ?? [];
    if (groupTextContentFingerprint(content) === groupChild.content) {
      continue;
    }
    const edits = editsByDrawing.get(group.drawing) ?? [];
    edits.push({ path: groupChild.path, contentXml: serializeTextBody(content) });
    editsByDrawing.set(group.drawing, edits);
  }

  const rewritten = new Map<RunContent, DrawingContent>();
  for (const [drawing, edits] of editsByDrawing) {
    const rawXml =
      drawing.rawXml === undefined ? undefined : replaceGroupTextBoxContent(drawing.rawXml, edits);
    if (rawXml !== undefined) {
      rewritten.set(drawing, { ...drawing, rawXml });
    }
  }

  return mapRuns(paragraph.content, (run) =>
    run.content.some((content) => absorbed.has(content) || rewritten.has(content))
      ? {
          ...run,
          content: run.content
            .filter((content) => !absorbed.has(content))
            .map((content) => rewritten.get(content) ?? content),
        }
      : run,
  );
};
