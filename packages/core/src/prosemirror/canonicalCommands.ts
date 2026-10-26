import type { Node as PMNode } from "prosemirror-model";
import type { Command, EditorState } from "prosemirror-state";

import type {
  FormattingPatch,
  ParagraphPropsPatch,
  RunPropsPatch,
} from "@stll/docx-core/ops";

import type {
  BlockContent,
  ColorValue,
  EmphasisMark,
  PreservedMarkup,
  ShadingProperties,
  TextEffect,
  UnderlineStyle,
} from "../types/document";
import { canonicalSelectionRanges } from "./canonicalSelectionRange";

/** Command meaning before canonical positions are resolved by the session. */
export type CanonicalCommandIntent =
  | { type: "generateTOC"; at: number; title: string }
  | {
      type: "setHyperlink";
      from: number;
      to: number;
      href: string;
      tooltip?: string;
    }
  | {
      type: "removeHyperlink";
      from: number;
      to: number;
      hyperlinkStyleId?: string;
    }
  | {
      type: "insertHyperlink";
      from: number;
      to: number;
      text: string;
      href: string;
      tooltip?: string;
    }
  | { type: "formatRun"; from: number; to: number; patch: RunPropsPatch }
  | { type: "formatParagraph"; at: number; patch: ParagraphPropsPatch }
  | { type: "toggleList"; kind: "bullet" | "decimal" }
  | { type: "changeListLevel"; direction: "increase" | "decrease" }
  | { type: "removeList" }
  | { type: "restartNumbering"; start?: number }
  | { type: "continueNumbering" }
  | {
      type: "setSectionBreak";
      breakType:
        "nextPage" | "nextColumn" | "continuous" | "evenPage" | "oddPage";
    }
  | { type: "removeSectionBreak" }
  | {
      type: "insertNote";
      at: number;
      kind: "footnote" | "endnote";
      id: number;
      content: BlockContent[];
    }
  | {
      type: "deleteNote";
      from: number;
      kind: "footnote" | "endnote";
      id: number;
    }
  | {
      type: "insertBreak";
      from: number;
      to: number;
      breakType: "page" | "textWrapping";
    };

type CanonicalCommandDescriptor = (
  state: EditorState,
) => readonly CanonicalCommandIntent[] | undefined;
const descriptors = new WeakMap<Command, CanonicalCommandDescriptor>();

/** Registration leaves execution and side-effect-free command probes unchanged. */
export const withCanonicalCommand = (
  command: Command,
  descriptor: CanonicalCommandDescriptor,
): Command => {
  descriptors.set(command, descriptor);
  return command;
};

export const getCanonicalCommandIntents = (
  command: Command,
  state: EditorState,
): readonly CanonicalCommandIntent[] | undefined =>
  descriptors.get(command)?.(state);

export const canonicalRunFormatting = (
  state: EditorState,
  patch: RunPropsPatch,
): {
  type: "formatRun";
  from: number;
  to: number;
  patch: FormattingPatch<{
    bold?: boolean | undefined;
    boldCs?: boolean | undefined;
    italic?: boolean | undefined;
    italicCs?: boolean | undefined;
    underline?:
      | {
          style: UnderlineStyle;
          color?: ColorValue;
        }
      | undefined;
    strike?: boolean | undefined;
    doubleStrike?: boolean | undefined;
    vertAlign?: "baseline" | "superscript" | "subscript" | undefined;
    smallCaps?: boolean | undefined;
    allCaps?: boolean | undefined;
    hidden?: boolean | undefined;
    noProof?: boolean | undefined;
    color?: ColorValue | undefined;
    highlight?:
      | "black"
      | "blue"
      | "cyan"
      | "darkBlue"
      | "darkCyan"
      | "darkGray"
      | "darkGreen"
      | "darkMagenta"
      | "darkRed"
      | "darkYellow"
      | "green"
      | "lightGray"
      | "magenta"
      | "none"
      | "red"
      | "white"
      | "yellow"
      | undefined;
    shading?: ShadingProperties | undefined;
    fontSize?: number | undefined;
    fontSizeCs?: number | undefined;
    fontFamily?:
      | {
          ascii?: string;
          hAnsi?: string;
          eastAsia?: string;
          cs?: string;
          hint?: "default" | "eastAsia" | "cs";
          asciiTheme?:
            | "majorAscii"
            | "majorHAnsi"
            | "majorEastAsia"
            | "majorBidi"
            | "minorAscii"
            | "minorHAnsi"
            | "minorEastAsia"
            | "minorBidi";
          hAnsiTheme?: string;
          eastAsiaTheme?: string;
          csTheme?: string;
        }
      | undefined;
    language?:
      | {
          val?: string;
          eastAsia?: string;
          bidi?: string;
        }
      | undefined;
    spacing?: number | undefined;
    position?: number | undefined;
    scale?: number | undefined;
    kerning?: number | undefined;
    effect?: TextEffect | undefined;
    emphasisMark?: EmphasisMark | undefined;
    emboss?: boolean | undefined;
    imprint?: boolean | undefined;
    outline?: boolean | undefined;
    shadow?: boolean | undefined;
    rtl?: boolean | undefined;
    cs?: boolean | undefined;
    styleId?: string | undefined;
    preserved?: PreservedMarkup | undefined;
  }>;
}[] =>
  canonicalSelectionRanges(state).map(({ from, to }) => ({
    type: "formatRun" as const,
    from,
    to,
    patch,
  }));

const canonicalParagraphFormatting = (
  state: EditorState,
  patch: ParagraphPropsPatch | ((paragraph: PMNode) => ParagraphPropsPatch),
) => {
  const intents: CanonicalCommandIntent[] = [];
  state.doc.nodesBetween(
    state.selection.from,
    state.selection.to,
    (node, position) => {
      if (node.type.name !== "paragraph") return;
      intents.push({
        type: "formatParagraph",
        at: position + 1,
        patch: typeof patch === "function" ? patch(node) : patch,
      });
    },
  );
  return intents;
};

export const withCanonicalParagraphFormatting = (
  command: Command,
  patch: ParagraphPropsPatch | ((paragraph: PMNode) => ParagraphPropsPatch),
): Command =>
  withCanonicalCommand(command, (state) =>
    canonicalParagraphFormatting(state, patch),
  );

/** Commands that read the selection start apply that same patch across selected paragraphs. */
export const withCanonicalStartParagraphFormatting = (
  command: Command,
  patch: (paragraph: PMNode) => ParagraphPropsPatch,
): Command =>
  withCanonicalCommand(command, (state) => {
    const paragraph = state.selection.$from.parent;
    if (paragraph.type.name !== "paragraph") return [];
    return canonicalParagraphFormatting(state, patch(paragraph));
  });
