import type { DocxEditorRef, FindResult } from "@stll/folio-vue";

export const buildFindNotificationBridge = (
  getEditor: () => DocxEditorRef | null,
  notifications: (FindResult | null)[],
) => ({
  find: (text: string) => getEditor()?.find(text),
  findNext: () => getEditor()?.findNext(),
  findPrevious: () => getEditor()?.findPrevious(),
  clearFind: () => getEditor()?.clearFind(),
  takeNotifications: () => notifications.splice(0),
});

export type FindNotificationBridge = ReturnType<typeof buildFindNotificationBridge>;
