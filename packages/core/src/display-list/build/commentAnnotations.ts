import { getParagraphText } from "../../docx/paragraphParser";
import type { Comment } from "../../types/document";

export type DisplayCommentInput = {
  readonly id: number;
  readonly author: string;
  readonly contents: string;
};

export const displayCommentsFrom = (
  comments: readonly Comment[],
): readonly DisplayCommentInput[] => {
  const inputs: DisplayCommentInput[] = [];
  for (const { id, author, content } of comments) {
    const paragraphs: string[] = [];
    for (const paragraph of content) {
      paragraphs.push(getParagraphText(paragraph));
    }
    inputs.push({ id, author, contents: paragraphs.join("\n") });
  }
  return inputs;
};
