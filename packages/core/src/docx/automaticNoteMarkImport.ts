import { PARSE_WARNING_CODES } from "@stll/docx-core/model";

import type { ParseContext } from "./parseContext";
import {
  getChildElements,
  getLocalName,
  getNamespaceUri,
  NAMESPACES,
  selectAlternateContentBranch,
  WORDPROCESSINGML_NAMESPACE_URIS,
  type XmlElement,
} from "./xmlParser";

type NoteKind = "footnote" | "endnote";

const misplacedMark = Symbol("misplaced automatic note mark");

type MarkNoteImportOptions = {
  root: XmlElement;
  note?: NoteKind;
  context: ParseContext | undefined;
};

type NoteMarkScope = { story: NoteKind | null; runParent: boolean };

/** Mark XML for verbatim capture by the run parser. The flag follows XML clones. */
export const markMisplacedAutomaticNoteMarks = ({
  root,
  note,
  context,
}: MarkNoteImportOptions): void => {
  const visit = (element: XmlElement, scope: NoteMarkScope): void => {
    const { story, runParent } = scope;
    const namespace = getNamespaceUri(element);
    const local = getLocalName(element.name);
    if (namespace === NAMESPACES.mc && local === "AlternateContent") {
      const branch = selectAlternateContentBranch(element);
      if (branch !== undefined)
        for (const child of getChildElements(branch)) visit(child, scope);
      return;
    }
    const word = WORDPROCESSINGML_NAMESPACE_URIS.has(namespace ?? "");
    if (
      word &&
      runParent &&
      (local === "footnoteRef" || local === "endnoteRef")
    ) {
      const kind = local === "footnoteRef" ? "footnote" : "endnote";
      if (kind !== story && !(misplacedMark in element)) {
        Object.assign(element, { [misplacedMark]: true });
        context?.warn({
          code: PARSE_WARNING_CODES.misplacedNoteMark,
          element: element.name ?? local,
        });
      }
      return;
    }
    // A drawing's text box is a story of its own even inside a note part.
    const children = getChildElements(element);
    if (children.length === 0) return;
    const childScope = {
      story: word && local === "txbxContent" ? null : story,
      runParent: word && local === "r",
    };
    for (const child of children) visit(child, childScope);
  };
  visit(root, { story: note ?? null, runParent: false });
};

export const isMisplacedAutomaticNoteMark = (element: XmlElement): boolean =>
  misplacedMark in element;
