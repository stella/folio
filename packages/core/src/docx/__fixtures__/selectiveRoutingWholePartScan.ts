/**
 * The selective save's paragraph routing as it read the serialized part:
 * scanned and parsed whole on every save, even when the canonical block
 * replay had written it as the source part with a few blocks replaced. Kept
 * only as the oracle for selectiveXmlPatchReplay.property.test.ts, which
 * checks that routing a replayed part from the regions it replaced writes the
 * same bytes.
 */
import { resolveParagraphIdentities } from "../paraIdAttribute";
import {
  isXmlNameBoundary,
  type ParagraphContainer,
  type ScannedParagraph,
  spliceSupportsRootNamespaces,
  spliceXml,
  type XmlSplice,
} from "../selectiveXmlPatch";
import {
  matchCloseTag,
  matchOpenTag,
  PARAGRAPH_SCAN_NAMES,
  resolveSelectiveScanPrefixes,
  splicesAsCanonical,
} from "../wordprocessingPrefixes";

type TrackedContainerKey =
  "fallback" | "alternateContent" | "textBox" | "tableCell";

/**
 * Every `<w:p>` element of `xml`, in the document order of its opening tags,
 * with the `w14:paraId` that tag carries and the containers around it.
 *
 * Document order is what makes the array an ordinal space: within one story
 * (see {@link ParagraphContainer}), the *n*th paragraph of the source part and
 * the *n*th paragraph of the model's serialization name the same paragraph,
 * which is the only way to address a paragraph the producer gave no id. A
 * `<w:p>` nested inside another (inside `mc:AlternateContent`, a text box) is
 * one entry of its own, exactly as {@link countParagraphElements} counts it.
 * An unterminated paragraph keeps `end === start`: it occupies its ordinal but
 * no splice can be built from it.
 */
function scanParagraphs(xml: string): ScannedParagraph[] {
  const resolution = resolveSelectiveScanPrefixes(xml);
  const mcPrefixes =
    resolution.type === "resolved" ? resolution.prefixes.mc : ["mc"];
  const paragraphs: ScannedParagraph[] = [];
  const open: number[] = [];
  const depth: Record<TrackedContainerKey, number> = {
    fallback: 0,
    alternateContent: 0,
    textBox: 0,
    tableCell: 0,
  };
  const containers = [
    { key: "fallback", prefixes: mcPrefixes, name: "Fallback" },
    { key: "alternateContent", prefixes: mcPrefixes, name: "AlternateContent" },
    { key: "textBox", prefixes: ["w"], name: "txbxContent" },
    { key: "tableCell", prefixes: ["w"], name: "tc" },
  ] as const;
  let pos = 0;

  scan: while (pos < xml.length) {
    const tagStart = xml.indexOf("<", pos);
    if (tagStart === -1) {
      break;
    }

    if (xml.startsWith("</w:p>", tagStart)) {
      const slot = open.pop();
      const end = tagStart + "</w:p>".length;
      const paragraph = slot === undefined ? undefined : paragraphs[slot];
      if (paragraph) {
        paragraph.end = end;
      }
      pos = end;
      continue;
    }

    for (const { key, prefixes, name } of containers) {
      const closeLength = matchCloseTag(xml, tagStart, prefixes, name);
      if (closeLength !== -1) {
        depth[key] = Math.max(0, depth[key] - 1);
        pos = tagStart + closeLength;
        continue scan;
      }
      if (matchOpenTag(xml, tagStart, prefixes, name) !== -1) {
        const tagEnd = xml.indexOf(">", tagStart);
        if (tagEnd === -1) {
          break scan;
        }
        if (xml[tagEnd - 1] !== "/") {
          depth[key] += 1;
        }
        pos = tagEnd + 1;
        continue scan;
      }
    }

    if (
      !xml.startsWith("<w:p", tagStart) ||
      !isXmlNameBoundary(xml[tagStart + 4])
    ) {
      pos = tagStart + 1;
      continue;
    }

    const tagEnd = xml.indexOf(">", tagStart);
    if (tagEnd === -1) {
      break;
    }
    const openTag = xml.slice(tagStart, tagEnd + 1);
    const paraId =
      /\bw14:paraId\s*=\s*(?<quote>["'])(?<id>[^"']+)\k<quote>/u.exec(openTag)
        ?.groups?.["id"];
    const container: ParagraphContainer = {
      inFallback: depth.fallback > 0,
      inAlternateContent: depth.alternateContent > 0,
      textBoxDepth: depth.textBox,
      tableDepth: depth.tableCell,
    };

    if (xml[tagEnd - 1] === "/") {
      // Self-closing <w:p ... /> — resolved immediately, never pushed.
      paragraphs.push({ start: tagStart, end: tagEnd + 1, paraId, container });
    } else {
      open.push(paragraphs.length);
      paragraphs.push({ start: tagStart, end: tagStart, paraId, container });
    }
    pos = tagEnd + 1;
  }

  return paragraphs;
}

/** Paragraph ids folio writes and this module may have to remove again. */
const MINTED_PARA_ID_ATTRIBUTE = /\sw14:(paraId|textId)="[^"]*"/gu;

/**
 * The replacement for a paragraph whose source open tag carries no paraId.
 *
 * The model always has a key, because folio mints one for every paragraph that
 * arrives without one, and the serializer writes whatever key the model holds.
 * Writing a minted key into the package would upgrade a producer's id-less
 * document to a partly-id'd one on any edit, declare a `w14` namespace the
 * part never had, and turn an id the minter itself calls positional into an
 * authored one the next reader trusts. So the save keeps the author's
 * convention: `ensureParaIds` is the one pass that gives a package ids, and it
 * runs when a host asks for it. An id the source *does* carry (a `w14:textId`
 * beside no paraId) is the author's and stays.
 */
const withoutMintedIds = (
  paragraphXml: string,
  sourceOpenTag: string,
): string => {
  const tagEnd = paragraphXml.indexOf(">");
  if (tagEnd === -1) {
    return paragraphXml;
  }
  const openTag = paragraphXml
    .slice(0, tagEnd + 1)
    .replace(MINTED_PARA_ID_ATTRIBUTE, (attribute, name: string) => {
      const sourceAttribute = new RegExp(
        `\\bw14:${name}\\s*=\\s*(["'])[^"']*\\1`,
        "u",
      );
      return sourceAttribute.test(sourceOpenTag) ? attribute : "";
    });
  return openTag + paragraphXml.slice(tagEnd + 1);
};

/** Where each changed paragraph's re-serialized XML goes, or why it cannot. */
type ParagraphRouting =
  { type: "routed"; xml: string } | { type: "refused"; reason: string };

/** An ordinal space of one part; see {@link ParagraphContainer}. */
type Story = "main" | "text-box";

const storyOf = ({ textBoxDepth }: ParagraphContainer): Story =>
  textBoxDepth > 0 ? "text-box" : "main";

const sameContainer = (a: ParagraphContainer, b: ParagraphContainer): boolean =>
  a.inAlternateContent === b.inAlternateContent &&
  a.textBoxDepth === b.textBoxDepth &&
  a.tableDepth === b.tableDepth;

/** The paragraphs of one part a splice may address, indexed the ways routing asks for them. */
type AddressableParagraphs = {
  stories: ReadonlyMap<Story, readonly ScannedParagraph[]>;
  /** Each paragraph's ordinal within its story. */
  ordinals: ReadonlyMap<ScannedParagraph, number>;
  byParaId: ReadonlyMap<string, readonly ScannedParagraph[]>;
};

const addressableParagraphs = (xml: string): AddressableParagraphs => {
  const stories = new Map<Story, ScannedParagraph[]>();
  const ordinals = new Map<ScannedParagraph, number>();
  const byParaId = new Map<string, ScannedParagraph[]>();
  for (const paragraph of scanParagraphs(xml)) {
    if (paragraph.container.inFallback) {
      continue;
    }
    const story = storyOf(paragraph.container);
    const sequence = stories.get(story) ?? [];
    ordinals.set(paragraph, sequence.length);
    sequence.push(paragraph);
    stories.set(story, sequence);
    if (paragraph.paraId !== undefined) {
      const named = byParaId.get(paragraph.paraId) ?? [];
      named.push(paragraph);
      byParaId.set(paragraph.paraId, named);
    }
  }
  return { stories, ordinals, byParaId };
};

/** The paraIds written on `<w:p>` elements inside `mc:Fallback`. */
const fallbackParaIds = (xml: string): ReadonlySet<string> =>
  new Set(
    scanParagraphs(xml)
      .filter(
        ({ container, paraId }) => container.inFallback && paraId !== undefined,
      )
      .map(({ paraId }) => paraId as string),
  );

/**
 * The prefix of the part's root element. The scan reads WordprocessingML by
 * its conventional `w:` prefix, so a part that binds it to another prefix has
 * no paragraphs the scan can see, and a spliced `w:` paragraph would not
 * resolve under its root.
 */
const rootElementName = (xml: string): string | undefined =>
  /<(?<name>[A-Za-z_][\w.-]*(?::[\w.-]+)?)/u.exec(xml)?.groups?.["name"];

/**
 * The id of the nearest paragraph on one side of `paragraph` in its story that
 * both parts name exactly once: the witness that places it in its story.
 */
const nearestSharedParaId = (
  own: AddressableParagraphs,
  other: AddressableParagraphs,
  paragraph: ScannedParagraph,
  step: -1 | 1,
): string | null => {
  const sequence = own.stories.get(storyOf(paragraph.container)) ?? [];
  for (
    let ordinal = (own.ordinals.get(paragraph) ?? -1) + step;
    ordinal >= 0 && ordinal < sequence.length;
    ordinal += step
  ) {
    const paraId = sequence[ordinal]?.paraId;
    if (
      paraId !== undefined &&
      own.byParaId.get(paraId)?.length === 1 &&
      other.byParaId.get(paraId)?.length === 1
    ) {
      return paraId;
    }
  }
  return null;
};

/**
 * Route every changed paragraph from the model's serialization to its region
 * of the source part.
 *
 * One owner for the question the patch turns on: *which paragraph of the file
 * is this model paragraph?* The answer is local to the paragraph. A paraId the
 * source writes once names it, robust to anything the model does elsewhere; a
 * paragraph the source never named is located by its ordinal within its story,
 * which is sound exactly when {@link resolveParagraphIdentities} says that
 * story lines up. Nothing else in the part has to agree: the model may drop a
 * text box it cannot read or write a Fallback it does not own, and an edit in
 * the main flow is still one paragraph of the main flow.
 *
 * What the routing still refuses is a paragraph it cannot place with
 * certainty, or one whose place changed:
 * - a paraId missing from, or written twice by, either part (among the
 *   paragraphs outside `mc:Fallback`);
 * - an id the source writes only inside `mc:Fallback`, which the model does
 *   not own;
 * - a paragraph inside `mc:AlternateContent`, whose Fallback repeats it and
 *   would contradict the edited Choice;
 * - a paragraph whose container (table-cell nesting, text box, alternate
 *   content) differs between the two parts;
 * - an authored paragraph whose nearest shared neighbours differ, i.e. that
 *   moved within its story;
 * - an id-less paragraph whose story's ordinals do not line up.
 */
const routeChangedParagraphs = (
  originalXml: string,
  serializedXml: string,
  changedIds: ReadonlySet<string>,
): ParagraphRouting => {
  // The scan reads `<w:p` / `w14:paraId` literally and the splice writes the
  // serializer's `w:` markup, so both hold only for a part spelled that way:
  // a root under another prefix, paragraphs or ids under a second
  // WordprocessingML prefix, `w14` under another prefix, or a nested
  // rebinding all refuse.
  if (!splicesAsCanonical(originalXml, PARAGRAPH_SCAN_NAMES)) {
    return {
      type: "refused",
      reason: `non-canonical-wordprocessingml-prefix: ${rootElementName(originalXml) ?? "unknown"}`,
    };
  }

  const original = addressableParagraphs(originalXml);
  const serialized = addressableParagraphs(serializedXml);
  let originalFallbackIds: ReadonlySet<string> | undefined;
  const storyAlignment = new Map<Story, boolean>();
  const storyAligned = (story: Story): boolean => {
    let aligned = storyAlignment.get(story);
    if (aligned === undefined) {
      aligned = resolveParagraphIdentities({
        sourceParaIds: (original.stories.get(story) ?? []).map(
          ({ paraId }) => paraId,
        ),
        serializedParaIds: (serialized.stories.get(story) ?? []).map(
          ({ paraId }) => paraId,
        ),
      }).ordinalsAligned;
      storyAlignment.set(story, aligned);
    }
    return aligned;
  };

  const splices: XmlSplice[] = [];
  for (const id of changedIds) {
    const inOriginal = original.byParaId.get(id) ?? [];
    const inSerialized = serialized.byParaId.get(id) ?? [];
    if (inOriginal.length > 1) {
      return { type: "refused", reason: `duplicate-paraId-in-original: ${id}` };
    }
    if (inOriginal.length === 1 && inSerialized.length === 0) {
      return {
        type: "refused",
        reason: `paraId-not-found-in-serialized: ${id}`,
      };
    }
    if (inSerialized.length === 0) {
      return { type: "refused", reason: `paraId-not-found-in-original: ${id}` };
    }
    if (inSerialized.length > 1) {
      return {
        type: "refused",
        reason: `duplicate-paraId-in-serialized: ${id}`,
      };
    }
    // SAFETY: inSerialized.length === 1 verified above
    const replacement = inSerialized[0]!;
    if (replacement.end <= replacement.start) {
      return { type: "refused", reason: `unterminated-paragraph: ${id}` };
    }
    const newXml = serializedXml.slice(replacement.start, replacement.end);

    const authored = inOriginal[0];
    let source: ScannedParagraph;
    if (authored) {
      source = authored;
    } else {
      originalFallbackIds ??= fallbackParaIds(originalXml);
      if (originalFallbackIds.has(id)) {
        return { type: "refused", reason: `paraId-only-in-fallback: ${id}` };
      }
      // The model minted this id: the paragraph is its ordinal in its story.
      const story = storyOf(replacement.container);
      if (!storyAligned(story)) {
        return {
          type: "refused",
          reason: `unaligned-paragraph-ordinals: ${id}`,
        };
      }
      const positional =
        original.stories.get(story)?.[
          serialized.ordinals.get(replacement) ?? -1
        ];
      if (!positional) {
        return {
          type: "refused",
          reason: `paraId-not-found-in-original: ${id}`,
        };
      }
      source = positional;
    }

    if (source.end <= source.start) {
      return { type: "refused", reason: `unterminated-paragraph: ${id}` };
    }
    if (source.container.inAlternateContent) {
      return {
        type: "refused",
        reason: `paragraph-in-alternate-content: ${id}`,
      };
    }
    if (!sameContainer(source.container, replacement.container)) {
      return { type: "refused", reason: `container-changed: ${id}` };
    }
    if (
      authored &&
      (nearestSharedParaId(original, serialized, source, -1) !==
        nearestSharedParaId(serialized, original, replacement, -1) ||
        nearestSharedParaId(original, serialized, source, 1) !==
          nearestSharedParaId(serialized, original, replacement, 1))
    ) {
      return { type: "refused", reason: `paragraph-moved: ${id}` };
    }

    const sourceXml = originalXml.slice(source.start, source.end);
    splices.push({
      start: source.start,
      end: source.end,
      newXml: authored ? newXml : withoutMintedIds(newXml, sourceXml),
    });
  }

  if (
    !spliceSupportsRootNamespaces({
      originalXml,
      serializedXml,
      fragments: splices.map(({ newXml }) => newXml),
    })
  ) {
    return { type: "refused", reason: "root-namespace-requirements-changed" };
  }

  const candidate = spliceXml(originalXml, splices);
  if (candidate === null) {
    return { type: "refused", reason: "unsafe-paragraph-splices" };
  }
  if (!splicesAsCanonical(candidate, PARAGRAPH_SCAN_NAMES)) {
    return { type: "refused", reason: "replacement-namespace-conflict" };
  }
  return { type: "routed", xml: candidate };
};

/**
 * Build a patched document.xml by splicing new paragraph XML into
 * the original at the correct offsets. Only changed paragraphs
 * are replaced; everything else is preserved byte-for-byte.
 *
 * Returns null when a changed paragraph cannot be routed (see
 * {@link routeChangedParagraphs}) or {@link spliceXml} refuses the result; the
 * caller then falls back to a full repack, whose parts are all re-serialized
 * from one model.
 */
export function buildPatchedDocumentXml(
  originalXml: string,
  serializedXml: string,
  changedIds: Set<string>,
): string | null {
  if (changedIds.size === 0) {
    return originalXml;
  }
  const routing = routeChangedParagraphs(
    originalXml,
    serializedXml,
    changedIds,
  );
  return routing.type === "refused" ? null : routing.xml;
}
