import type {
  DocumentBody,
  Endnote,
  Footnote,
  HeaderFooter,
  Hyperlink,
  ParagraphContent,
  Run,
} from "../types/document";
import { visitDocxParagraphs } from "./paragraphTraversal";

type NormalizeRenderedPageBreakHintsInput = {
  documentBody: DocumentBody;
  headers?: Map<string, HeaderFooter>;
  footers?: Map<string, HeaderFooter>;
  footnotes?: readonly Footnote[];
  endnotes?: readonly Endnote[];
};

export const normalizeRenderedPageBreakHints = ({
  documentBody,
  headers,
  footers,
  footnotes,
  endnotes,
}: NormalizeRenderedPageBreakHintsInput): void => {
  visitDocxParagraphs({ documentBody, headers, footers, footnotes, endnotes }, (paragraph) => {
    if (
      paragraph.renderedPageBreakBefore === true &&
      !retainsLeadingRenderedPageBreak(paragraph.content)
    ) {
      delete paragraph.renderedPageBreakBefore;
    }
  });
};

const retainsLeadingRenderedPageBreak = (content: readonly ParagraphContent[]): boolean => {
  const scan = { sawRenderedPageBreak: false };
  for (const item of content) {
    const result = scanParagraphContent(item, scan);
    if (result !== undefined) {
      return result;
    }
  }
  return false;
};

type RenderedPageBreakScan = {
  sawRenderedPageBreak: boolean;
};

const scanRun = (run: Run, scan: RenderedPageBreakScan): boolean | undefined => {
  for (const content of run.content) {
    if (content.type === "renderedPageBreak") {
      scan.sawRenderedPageBreak = true;
      continue;
    }
    return scan.sawRenderedPageBreak;
  }
  return undefined;
};

const scanHyperlink = (hyperlink: Hyperlink, scan: RenderedPageBreakScan): boolean | undefined => {
  for (const child of hyperlink.children) {
    if (child.type !== "run") {
      continue;
    }
    const result = scanRun(child, scan);
    if (result !== undefined) {
      return result;
    }
  }
  return undefined;
};

const scanInlineContent = (
  content: readonly (Run | Hyperlink)[],
  scan: RenderedPageBreakScan,
): boolean | undefined => {
  for (const child of content) {
    const result = child.type === "run" ? scanRun(child, scan) : scanHyperlink(child, scan);
    if (result !== undefined) {
      return result;
    }
  }
  return undefined;
};

const scanParagraphContent = (
  content: ParagraphContent,
  scan: RenderedPageBreakScan,
): boolean | undefined => {
  if (content.type === "run") {
    return scanRun(content, scan);
  }
  if (content.type === "hyperlink") {
    return scanHyperlink(content, scan);
  }
  if (content.type === "simpleField") {
    return scanInlineContent(content.content, scan);
  }
  if (
    content.type === "insertion" ||
    content.type === "deletion" ||
    content.type === "moveFrom" ||
    content.type === "moveTo"
  ) {
    return scanInlineContent(content.content, scan);
  }
  if (content.type === "inlineSdt") {
    for (const child of content.content) {
      const result = scanParagraphContent(child, scan);
      if (result !== undefined) {
        return result;
      }
    }
    return undefined;
  }
  if (content.type === "complexField") {
    // The serializer always emits a leading w:fldChar for a complex field.
    return scan.sawRenderedPageBreak;
  }
  // Range markers and math are not visible to the raw leading-break detector.
  return undefined;
};
