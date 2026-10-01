/** Stable package-story addressing, shared by operations and editor projections. */
import { panic } from "better-result";
import type { Document, DocumentBody, HeaderFooter } from "../model/document";
import { OP_STORIES, type OpStory } from "./types";

export const sameStory = (left: OpStory, right: OpStory): boolean => {
  if (left === OP_STORIES.MAIN || right === OP_STORIES.MAIN) return left === right;
  if (left.kind !== right.kind) return false;
  switch (left.kind) {
    case "header":
    case "footer":
      return (right.kind === "header" || right.kind === "footer") && left.rId === right.rId;
    case "footnote":
    case "endnote":
      return (right.kind === "footnote" || right.kind === "endnote") && left.id === right.id;
    default: {
      const unreachable: never = left;
      return unreachable;
    }
  }
};

export const findStoryBody = (document: Document, story: OpStory): DocumentBody | undefined => {
  if (story === OP_STORIES.MAIN) return document.package.document;
  switch (story.kind) {
    case "header":
      return document.package.headers?.get(story.rId);
    case "footer":
      return document.package.footers?.get(story.rId);
    case "footnote":
      return document.package.footnotes?.find(({ id }) => id === story.id);
    case "endnote":
      return document.package.endnotes?.find(({ id }) => id === story.id);
    default: {
      const unreachable: never = story;
      return unreachable;
    }
  }
};

export const storyBody = (document: Document, story: OpStory): DocumentBody =>
  findStoryBody(document, story) ?? panic("An operation must address an existing document story.");

export const documentStories = (document: Document): OpStory[] => {
  const out: OpStory[] = [OP_STORIES.MAIN];
  for (const rId of document.package.headers?.keys() ?? []) out.push({ kind: "header", rId });
  for (const rId of document.package.footers?.keys() ?? []) out.push({ kind: "footer", rId });
  for (const { id } of document.package.footnotes ?? []) out.push({ kind: "footnote", id });
  for (const { id } of document.package.endnotes ?? []) out.push({ kind: "endnote", id });
  return out;
};

type ReplaceStoryBodyOptions = { document: Document; story: OpStory; body: DocumentBody };
export const replaceStoryBody = ({ document, story, body }: ReplaceStoryBodyOptions): Document => {
  if (story === OP_STORIES.MAIN)
    return { ...document, package: { ...document.package, document: body } };
  const pkg = document.package;
  switch (story.kind) {
    case "header":
    case "footer": {
      const parts = story.kind === "header" ? pkg.headers : pkg.footers;
      const previous = parts?.get(story.rId);
      if (!previous) return panic("Replacing a header/footer requires an existing part.");
      const next: HeaderFooter = { ...previous, content: body.content };
      const replaced = new Map(parts);
      replaced.set(story.rId, next);
      let sectionsChanged = false;
      const sections = pkg.document.sections?.map((section) => {
        const bound = story.kind === "header" ? section.headers : section.footers;
        const references =
          story.kind === "header"
            ? section.properties.headerReferences
            : section.properties.footerReferences;
        const variants = references?.filter(({ rId }) => rId === story.rId) ?? [];
        if (!bound || variants.length === 0) return section;
        sectionsChanged = true;
        const updated = new Map(bound);
        for (const { type } of variants) updated.set(type, next);
        return story.kind === "header"
          ? { ...section, headers: updated }
          : { ...section, footers: updated };
      });
      const main = sectionsChanged && sections ? { ...pkg.document, sections } : pkg.document;
      return story.kind === "header"
        ? { ...document, package: { ...pkg, document: main, headers: replaced } }
        : { ...document, package: { ...pkg, document: main, footers: replaced } };
    }
    case "footnote": {
      const footnotes = pkg.footnotes?.map((note) =>
        note.id === story.id ? { ...note, content: body.content } : note,
      );
      if (!footnotes) return panic("Replacing a footnote requires its collection.");
      return { ...document, package: { ...pkg, footnotes } };
    }
    case "endnote": {
      const endnotes = pkg.endnotes?.map((note) =>
        note.id === story.id ? { ...note, content: body.content } : note,
      );
      if (!endnotes) return panic("Replacing an endnote requires its collection.");
      return { ...document, package: { ...pkg, endnotes } };
    }
    default: {
      const unreachable: never = story;
      return unreachable;
    }
  }
};
