import type { FlowBlock } from "./types";

export const continuesNumberedSequence = (
  previous: FlowBlock | undefined,
  current: FlowBlock,
): boolean => {
  if (previous?.kind !== "paragraph" || current.kind !== "paragraph") {
    return false;
  }
  const previousNumPr = previous.attrs?.numPr;
  const currentNumPr = current.attrs?.numPr;
  if (previousNumPr?.numId === undefined || currentNumPr?.numId === undefined) {
    return false;
  }
  return previousNumPr.numId === currentNumPr.numId && previousNumPr.ilvl === currentNumPr.ilvl;
};

export const continuesTabbedParagraphSequence = (
  previous: FlowBlock | undefined,
  current: FlowBlock,
): boolean => {
  if (previous?.kind !== "paragraph" || current.kind !== "paragraph") {
    return false;
  }
  if (!previous.attrs?.styleId || previous.attrs.styleId !== current.attrs?.styleId) {
    return false;
  }
  if (
    !previous.runs.some((run) => run.kind === "tab") ||
    !current.runs.some((run) => run.kind === "tab")
  ) {
    return false;
  }
  const previousIndent = previous.attrs.indent;
  const currentIndent = current.attrs.indent;
  if (
    previousIndent?.left !== currentIndent?.left ||
    previousIndent?.right !== currentIndent?.right ||
    previousIndent?.firstLine !== currentIndent?.firstLine ||
    previousIndent?.hanging !== currentIndent?.hanging
  ) {
    return false;
  }
  const previousTabs = previous.attrs.tabs ?? [];
  const currentTabs = current.attrs.tabs ?? [];
  return (
    previousTabs.length === currentTabs.length &&
    previousTabs.every((tab, index) => {
      const currentTab = currentTabs[index];
      if (!currentTab) {
        return false;
      }
      return (
        tab.val === currentTab.val && tab.pos === currentTab.pos && tab.leader === currentTab.leader
      );
    })
  );
};
