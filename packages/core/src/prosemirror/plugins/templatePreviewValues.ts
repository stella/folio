/**
 * Template Preview Values Plugin
 *
 * Live fill preview for template documents: given a map of field path →
 * typed value, every matching `{{path}}` marker is visually replaced by
 * the value — the marker text is hidden by an inline decoration and the
 * value is injected as a widget in its place. The document itself is
 * never modified; clearing the preview restores the markers.
 *
 * The same preview hides conditional blocks the host reports as not
 * applying: `conditions` maps an `{% if expr %}` expression to whether
 * its block applies, and a `false` block is hidden from its opener
 * through its `{% endif %}`. folio evaluates no expression.
 *
 * Two render modes: `highlighted` paints the substituted values with the
 * preview accent so it is unmistakably a preview; `plain` renders them
 * as ordinary text, approximating the final filled document.
 *
 * Updates are pushed via {@link setTemplatePreviewValues}; the host wires
 * this to its fill inputs.
 */

import type { Node as PMNode } from "prosemirror-model";
import { Plugin, PluginKey } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { Decoration, DecorationSet } from "prosemirror-view";

import type { DirectiveRange } from "./templateDirectives";
import { scanDirectives } from "./templateDirectives";

/** One formatted span of a rich preview value. The flags layer over the
 *  marker's host formatting: a host-bold marker keeps its bold, a span's
 *  own bold/italic ORs in on top. */
export type TemplatePreviewSpan = {
  text: string;
  bold?: boolean;
  italic?: boolean;
};

/** A preview value: plain text, or formatted spans when the value carries
 *  its own bold/italic (e.g. a registry lookup's formatted rendering). */
export type TemplatePreviewValue = string | { runs: TemplatePreviewSpan[] };

/** Concatenated plain text of a preview value. */
export const templatePreviewValueText = (value: TemplatePreviewValue): string =>
  typeof value === "string" ? value : value.runs.map((run) => run.text).join("");

/** Stable identity of a preview value (text + formatting), for decoration
 *  keys and change detection. Distinguishes a plain string from a rich
 *  value with the same text. */
export const templatePreviewValueFingerprint = (value: TemplatePreviewValue): string => {
  if (typeof value === "string") {
    return value;
  }
  return value.runs
    .map((run) => `${run.bold === true ? "b" : ""}${run.italic === true ? "i" : ""}:${run.text}`)
    .join("\u0000");
};

export type TemplatePreviewValues = {
  /** Field path → value to display in place of `{{path}}` markers. */
  values: Record<string, TemplatePreviewValue>;
  /** `highlighted` marks substitutions with the preview accent. */
  mode: "highlighted" | "plain";
  /**
   * `if` expression → whether its block applies. The key is the expression
   * exactly as written between `{% if` and `%}`, trimmed. `false` hides the
   * block; `true`, and an expression the map does not mention, leave it as
   * authored. The host decides truthiness: folio evaluates nothing.
   */
  conditions?: Record<string, boolean>;
};

/**
 * One `{{path}}` marker matched by an active preview value. Exposed on
 * the plugin state so the paged editor (whose visible pages never see
 * PM decorations) can project the same substitutions onto its overlay.
 */
export type TemplatePreviewEntry = {
  /** Inclusive PM doc position of the marker start. */
  from: number;
  /** Exclusive PM doc position of the marker end. */
  to: number;
  /** Field path the marker resolves to. */
  expr: string;
  /** The typed value displayed in place of the marker. */
  value: TemplatePreviewValue;
};

/**
 * One `{% if expr %}` … `{% endif %}` span hidden by a `false` condition,
 * exposed on the plugin state beside {@link TemplatePreviewEntry} so a
 * surface that never sees PM decorations can drop the same span.
 */
export type TemplatePreviewHiddenRange = {
  /** Inclusive PM doc position of the `{% if %}` opener start. */
  from: number;
  /** Exclusive PM doc position of the `{% endif %}` closer end. */
  to: number;
  /** The `if` expression, as written between `{% if` and `%}`, trimmed. */
  expr: string;
};

type TemplatePreviewState = {
  preview: TemplatePreviewValues | null;
  entries: TemplatePreviewEntry[];
  hidden: TemplatePreviewHiddenRange[];
  decorationSet: DecorationSet;
};

/** Hides the marker text a preview value stands in for. */
const ORIGINAL_CLASS = "folio-template-preview-original";
/** Hides a conditional block the host reported as not applying. */
const HIDDEN_CLASS = "folio-template-preview-hidden";

// Pin the PluginKey to a process-wide symbol so every module
// evaluation (Vite dev double-serve, @stll/folio re-export) resolves
// to the same key instance — otherwise host key-based lookups break.
const KEY_HOLDER_SYMBOL = Symbol.for("stll.folio.templatePreviewValuesKey");
type KeyHolder = {
  [KEY_HOLDER_SYMBOL]?: PluginKey<TemplatePreviewState>;
};
const keyHolder = globalThis as unknown as KeyHolder;
export const templatePreviewValuesKey: PluginKey<TemplatePreviewState> =
  keyHolder[KEY_HOLDER_SYMBOL] ??
  (keyHolder[KEY_HOLDER_SYMBOL] = new PluginKey<TemplatePreviewState>("templatePreviewValues"));

/** A rich span as DOM: text wrapped in `<em>` / `<strong>` per its flags. */
function buildSpanNode(span: TemplatePreviewSpan): Node {
  let node: Node = document.createTextNode(span.text);
  if (span.italic === true) {
    const em = document.createElement("em");
    em.append(node);
    node = em;
  }
  if (span.bold === true) {
    const strong = document.createElement("strong");
    strong.append(node);
    node = strong;
  }
  return node;
}

function buildValueWidget(
  value: TemplatePreviewValue,
  mode: TemplatePreviewValues["mode"],
): () => HTMLElement {
  return () => {
    const span = document.createElement("span");
    span.className =
      mode === "highlighted"
        ? "folio-template-preview-value folio-template-preview-value--highlighted"
        : "folio-template-preview-value";
    span.contentEditable = "false";
    if (typeof value === "string") {
      span.textContent = value;
      return span;
    }
    for (const run of value.runs) {
      span.append(buildSpanNode(run));
    }
    return span;
  };
}

/**
 * The `{% if %}` … `{% endif %}` spans a `false` condition hides, in document
 * order and without their nested hidden blocks: hiding a block already hides
 * everything inside it.
 *
 * Openers and closers pair off a kind-aware stack — `{% endif %}` closes the
 * nearest open `{% if %}`, `{% endfor %}` the nearest `{% for %}` — so a
 * mid-edit template with an unpaired opener or a stray closer hides nothing
 * rather than guessing a span. Inline markers pair like block ones: the fill
 * engine resolves an inline conditional within its paragraph, so the preview
 * follows. An `{% elif %}` or `{% else %}` branch sits inside the span and is
 * hidden with it.
 */
function collectHiddenRanges(
  ranges: readonly DirectiveRange[],
  conditions: Record<string, boolean> | undefined,
): TemplatePreviewHiddenRange[] {
  if (conditions === undefined) {
    return [];
  }

  const openers: DirectiveRange[] = [];
  const paired: TemplatePreviewHiddenRange[] = [];
  for (const range of [...ranges].sort((a, b) => a.from - b.from)) {
    if (range.kind === "if" || range.kind === "for") {
      openers.push(range);
      continue;
    }
    if (range.kind !== "endif" && range.kind !== "endfor") {
      continue;
    }
    const wanted = range.kind === "endif" ? "if" : "for";
    let opener: DirectiveRange | undefined;
    for (let index = openers.length - 1; index >= 0; index -= 1) {
      const candidate = openers[index];
      if (candidate?.kind === wanted) {
        opener = candidate;
        // Consume the matched opener and drop anything still open above it.
        openers.length = index;
        break;
      }
    }
    if (opener === undefined || opener.kind !== "if" || conditions[opener.expr] !== false) {
      continue;
    }
    paired.push({ from: opener.from, to: range.to, expr: opener.expr });
  }

  // Closers resolve inside-out, so order by document position and keep only
  // the outermost span of a nested pair.
  paired.sort((a, b) => a.from - b.from || b.to - a.to);
  const outermost: TemplatePreviewHiddenRange[] = [];
  for (const range of paired) {
    const previous = outermost.at(-1);
    if (previous && range.to <= previous.to) {
      continue;
    }
    outermost.push(range);
  }
  return outermost;
}

/** Everything the preview derives from the doc, off a single directive scan. */
type TemplatePreviewProjection = {
  entries: TemplatePreviewEntry[];
  hidden: TemplatePreviewHiddenRange[];
};

function collectPreviewEntries(
  ranges: readonly DirectiveRange[],
  values: Record<string, TemplatePreviewValue>,
  hidden: readonly TemplatePreviewHiddenRange[],
): TemplatePreviewEntry[] {
  const entries: TemplatePreviewEntry[] = [];
  for (const range of ranges) {
    // Field placeholders (`{{ path }}`) key the preview map by their path;
    // clause slots (`{{ clause("Name") }}`) key it by the slot name, which
    // `scanDirectives` exposes as the clause range's `expr`. The host
    // supplies the resolved clause text under that same slot name.
    if (range.kind !== "placeholder" && range.kind !== "clause") {
      continue;
    }
    const value = values[range.expr];
    if (value === undefined || templatePreviewValueText(value) === "") {
      continue;
    }
    // A marker inside a hidden block is not substituted: the value widget
    // sits beside the marker text rather than inside the hiding decoration,
    // so it would otherwise outlive its block.
    if (hidden.some((block) => range.from >= block.from && range.from < block.to)) {
      continue;
    }
    entries.push({ from: range.from, to: range.to, expr: range.expr, value });
  }
  return entries;
}

function projectPreview(
  doc: PMNode,
  preview: TemplatePreviewValues | null,
): TemplatePreviewProjection {
  if (!preview) {
    return { entries: [], hidden: [] };
  }
  const hasValues = Object.keys(preview.values).length > 0;
  const hasConditions =
    preview.conditions !== undefined && Object.keys(preview.conditions).length > 0;
  if (!hasValues && !hasConditions) {
    return { entries: [], hidden: [] };
  }

  const ranges = scanDirectives(doc);
  const hidden = collectHiddenRanges(ranges, preview.conditions);
  return {
    entries: hasValues ? collectPreviewEntries(ranges, preview.values, hidden) : [],
    hidden,
  };
}

/**
 * Hide one span: a block the span covers whole is hidden as a node, leaving no
 * empty line behind; a block it covers in part is hidden over the covered text
 * only, so an inline `{% if %}` in a paragraph of running text takes its own
 * span and nothing else.
 */
function pushHiddenDecorations(
  doc: PMNode,
  { from, to }: TemplatePreviewHiddenRange,
  out: Decoration[],
): void {
  doc.nodesBetween(from, to, (node, pos) => {
    if (!node.isBlock) {
      return false;
    }
    const contentFrom = pos + 1;
    const contentTo = pos + node.nodeSize - 1;
    if (from <= contentFrom && to >= contentTo) {
      out.push(Decoration.node(pos, pos + node.nodeSize, { class: HIDDEN_CLASS }));
      return false;
    }
    if (!node.isTextblock) {
      return true;
    }
    const spanFrom = Math.max(from, contentFrom);
    const spanTo = Math.min(to, contentTo);
    if (spanTo > spanFrom) {
      out.push(
        Decoration.inline(
          spanFrom,
          spanTo,
          { class: HIDDEN_CLASS },
          { inclusiveStart: false, inclusiveEnd: false },
        ),
      );
    }
    return false;
  });
}

function buildDecorationSet(
  doc: PMNode,
  { entries, hidden }: TemplatePreviewProjection,
  mode: TemplatePreviewValues["mode"],
): DecorationSet {
  if (entries.length === 0 && hidden.length === 0) {
    return DecorationSet.empty;
  }

  const decorations: Decoration[] = [];
  for (const range of hidden) {
    pushHiddenDecorations(doc, range, decorations);
  }
  for (const entry of entries) {
    decorations.push(
      Decoration.inline(
        entry.from,
        entry.to,
        { class: ORIGINAL_CLASS },
        { inclusiveStart: false, inclusiveEnd: false },
      ),
      Decoration.widget(entry.from, buildValueWidget(entry.value, mode), {
        side: 1,
        marks: [],
        ignoreSelection: true,
        key: `folio-template-preview-${entry.expr}-${entry.from}-${mode}-${templatePreviewValueFingerprint(entry.value)}`,
      }),
    );
  }
  return DecorationSet.create(doc, decorations);
}

export function createTemplatePreviewValuesPlugin(): Plugin<TemplatePreviewState> {
  return new Plugin<TemplatePreviewState>({
    key: templatePreviewValuesKey,
    state: {
      init(): TemplatePreviewState {
        return {
          preview: null,
          entries: [],
          hidden: [],
          decorationSet: DecorationSet.empty,
        };
      },
      apply(tr, prev, _oldState, newState): TemplatePreviewState {
        const meta = tr.getMeta(templatePreviewValuesKey) as
          | { preview: TemplatePreviewValues | null }
          | undefined;
        if (meta !== undefined) {
          const projection = projectPreview(newState.doc, meta.preview);
          return {
            preview: meta.preview,
            entries: projection.entries,
            hidden: projection.hidden,
            decorationSet: buildDecorationSet(
              newState.doc,
              projection,
              meta.preview?.mode ?? "plain",
            ),
          };
        }
        if (tr.docChanged && prev.preview) {
          const projection = projectPreview(newState.doc, prev.preview);
          return {
            preview: prev.preview,
            entries: projection.entries,
            hidden: projection.hidden,
            decorationSet: buildDecorationSet(newState.doc, projection, prev.preview.mode),
          };
        }
        return prev;
      },
    },
    props: {
      decorations(state) {
        return templatePreviewValuesKey.getState(state)?.decorationSet;
      },
    },
  });
}

/** Push (or clear, with `null`) the live fill preview into the editor. */
export const setTemplatePreviewValues = (
  view: EditorView,
  preview: TemplatePreviewValues | null,
): void => {
  view.dispatch(view.state.tr.setMeta(templatePreviewValuesKey, { preview }));
};
