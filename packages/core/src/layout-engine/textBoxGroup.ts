import type { TextBoxBlock } from "./types";

const textBoxGroupIds = new WeakMap<TextBoxBlock, string>();

export const setTextBoxGroupId = (textBox: TextBoxBlock, groupId: string): void => {
  textBoxGroupIds.set(textBox, groupId);
};

export const getTextBoxGroupId = (textBox: TextBoxBlock): string | undefined =>
  textBoxGroupIds.get(textBox);
