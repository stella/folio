import { panic, Result } from "better-result";

import { mintRevisionId } from "../../plugins/revisionIds";
import type { TrackedRevisionAncestor } from "../../schema/marks";

type ClipboardRevision = {
  type: "insertion" | "deletion";
  revisionId: number;
  author: string;
  date: string | null;
};

const CLIPBOARD_REVISION_ATTRIBUTES = [
  "id",
  "lang",
  "dir",
  "title",
  "style",
] as const;

export const authorFromClipboardRevision = (cite: string | null): string => {
  if (!cite) return "";
  const address = cite.startsWith("mailto:")
    ? cite.slice("mailto:".length)
    : cite;
  const decoded = Result.try(() => decodeURIComponent(address));
  return decoded.isOk() ? decoded.value : address;
};

export const prepareClipboardTrackedRevisions = (root: HTMLElement): void => {
  const elements = Array.from(root.querySelectorAll("ins, del"));
  const revisionsByElement = new Map<Element, ClipboardRevision>();
  for (const element of elements) {
    const type =
      element.tagName.toLowerCase() === "ins" ? "insertion" : "deletion";
    revisionsByElement.set(element, {
      type,
      revisionId: mintRevisionId(),
      author: authorFromClipboardRevision(element.getAttribute("cite")),
      date: element.getAttribute("datetime"),
    });
  }

  const ancestorsByElement = new Map<Element, TrackedRevisionAncestor[]>();
  for (const element of elements) {
    const revision = revisionsByElement.get(element);
    if (revision === undefined) {
      return panic("Clipboard revision map lost a parsed element");
    }
    const ancestors: TrackedRevisionAncestor[] = [];
    let parent = element.parentElement;
    while (parent !== null && parent !== root) {
      const parentTag = parent.tagName.toLowerCase();
      const ancestor = revisionsByElement.get(parent);
      if (parentTag === "ins" || parentTag === "del") {
        if (ancestor === undefined) {
          return panic("Clipboard revision map lost an ancestor element");
        }
        ancestors.push({
          type: ancestor.type,
          revisionId: ancestor.revisionId,
          author: ancestor.author,
          ...(ancestor.date === null ? {} : { date: ancestor.date }),
          outerWrapperCount: 0,
        });
      }
      parent = parent.parentElement;
    }
    ancestorsByElement.set(element, ancestors.toReversed());
  }

  for (const element of elements) {
    const revision = revisionsByElement.get(element);
    const ancestors = ancestorsByElement.get(element);
    if (revision === undefined || ancestors === undefined) {
      return panic("Clipboard revision preparation lost an element");
    }
    const span = root.ownerDocument.createElement("span");
    for (const attribute of CLIPBOARD_REVISION_ATTRIBUTES) {
      const value = element.getAttribute(attribute);
      if (value !== null) span.setAttribute(attribute, value);
    }
    const classes = (element.getAttribute("class") ?? "")
      .split(/\s+/u)
      .filter(
        (className) =>
          className.length > 0 &&
          className !== "docx-insertion" &&
          className !== "docx-deletion",
      );
    classes.push(`docx-${revision.type}`);
    span.className = classes.join(" ");
    span.dataset["revisionId"] = String(revision.revisionId);
    span.dataset["author"] = revision.author;
    if (revision.date !== null) span.dataset["date"] = revision.date;
    if (ancestors.length > 0) {
      span.dataset["revisionAncestors"] = JSON.stringify(ancestors);
    }
    span.append(...Array.from(element.childNodes));
    element.replaceWith(span);
  }
};
