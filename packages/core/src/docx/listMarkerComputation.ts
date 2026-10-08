/** Marker counters shared by parsing and canonical model normalization. */
import type { Paragraph } from "../types/document";
import type { NumberingMap } from "./numberingParser";
import { isNumberingReference } from "./numberingReference";
import { bulletMarkerFontName, convertBulletToUnicode } from "./bulletMarkers";
import { formatOoxmlCounter } from "./ooxmlCounterFormatter";

export type PreviousListState = {
  abstractNumId: number | null;
  fromStyle: boolean;
  numId: number | null;
};

export type ComputeListMarkerOptions = {
  numbering: NumberingMap | null;
  listCounters: Map<number, number[]>;
  abstractCounters: Map<number, number[]>;
  restartedNumIds: Set<number>;
  previousList: PreviousListState;
};

export const computeListMarker = (
  paragraph: Paragraph,
  {
    numbering,
    listCounters,
    abstractCounters,
    restartedNumIds,
    previousList,
  }: ComputeListMarkerOptions,
): void => {
  const listRendering = paragraph.listRendering;
  if (!listRendering || !numbering) {
    previousList.abstractNumId = null;
    previousList.fromStyle = false;
    previousList.numId = null;
    return;
  }

  const { numId, level } = listRendering;
  if (!isNumberingReference(numId)) {
    previousList.abstractNumId = null;
    previousList.fromStyle = false;
    previousList.numId = null;
    return;
  }

  const firstEncounter = !listCounters.has(numId);
  if (firstEncounter) {
    listCounters.set(numId, Array.from<number>({ length: 9 }).fill(Number.NaN));
  }

  let counters = listCounters.get(numId);
  if (!counters) {
    return;
  }

  const abstractNumId = numbering.getAbstractNumId(numId);
  const styleNumbering = paragraph.formatting?.numPrFromStyle;
  const resumesRestartedInstance =
    firstEncounter &&
    listRendering.startOverride === undefined &&
    abstractNumId !== null &&
    previousList.abstractNumId === abstractNumId &&
    previousList.numId !== null &&
    previousList.numId !== numId &&
    restartedNumIds.has(previousList.numId);
  const resumesStyleInstance =
    firstEncounter &&
    !styleNumbering &&
    listRendering.startOverride === undefined &&
    abstractNumId !== null &&
    previousList.abstractNumId === abstractNumId &&
    previousList.fromStyle === true;
  if (
    abstractNumId !== null &&
    (styleNumbering || resumesRestartedInstance || resumesStyleInstance)
  ) {
    const latestAbstractCounters = abstractCounters.get(abstractNumId);
    if (latestAbstractCounters) {
      // A paragraph whose numbering comes only from its style resumes the
      // latest compatible list instance. Word does this when an attachment
      // starts a fresh w:num (with a startOverride) and later paragraphs fall
      // back to the style's original w:num: the style continues the attachment
      // sequence instead of reviving its stale counters from earlier content.
      counters = latestAbstractCounters;
      listCounters.set(numId, counters);
    }
  }
  if (
    listRendering.startOverride !== undefined ||
    resumesRestartedInstance ||
    (previousList.numId === numId && restartedNumIds.has(numId))
  ) {
    restartedNumIds.add(numId);
  }
  if (abstractNumId !== null && level > 0) {
    const latestAbstractCounters = abstractCounters.get(abstractNumId);
    const missingParentCounters = counters.slice(0, level).every(Number.isNaN);
    if (missingParentCounters) {
      for (let i = 0; i < level; i += 1) {
        const latestCounter = latestAbstractCounters?.[i];
        counters[i] =
          latestCounter !== undefined && !Number.isNaN(latestCounter)
            ? latestCounter
            : (numbering.getLevel(numId, i)?.start ?? 1);
      }
    }
  }

  if (Number.isNaN(counters[level])) {
    counters[level] = (numbering.getLevel(numId, level)?.start ?? 1) - 1;
  }
  counters[level] = (counters[level] ?? 0) + 1;

  for (let i = level + 1; i < counters.length; i += 1) {
    counters[i] = Number.NaN;
  }

  // Word's default LISTNUM field advances the counter at one ilvl deeper
  // than the host paragraph. Mirror the toFlowBlocks logic here so the
  // marker substituted at parse time agrees with the renderer's counters —
  // otherwise a follow-up paragraph at that depth picks up the stale,
  // pre-substituted "(a)" instead of "(b)".
  const childAdvances = listRendering.implicitChildLevelAdvances ?? 0;
  if (childAdvances > 0 && level + 1 < counters.length) {
    const childCounter = counters[level + 1];
    counters[level + 1] =
      (childCounter === undefined || Number.isNaN(childCounter) ? 0 : childCounter) + childAdvances;
  }

  if (abstractNumId !== null) {
    abstractCounters.set(abstractNumId, counters);
  }
  previousList.abstractNumId = abstractNumId;
  previousList.fromStyle = Boolean(styleNumbering);
  previousList.numId = numId;

  const pattern = listRendering.marker;

  if (listRendering.isBullet) {
    listRendering.marker = convertBulletToUnicode(
      pattern || "",
      bulletMarkerFontName(listRendering.markerFormatting),
    );
    previousList.abstractNumId = null;
    previousList.fromStyle = false;
    previousList.numId = null;
    return;
  }

  let computedMarker = pattern;
  const currentLevelInfo = numbering.getLevel(numId, level);
  const useLegalNumbering = currentLevelInfo?.isLgl === true || listRendering.isLegal === true;

  for (let lvl = 0; lvl <= level; lvl += 1) {
    const placeholder = `%${lvl + 1}`;
    if (computedMarker.includes(placeholder)) {
      const value = counters[lvl] ?? 0;
      const levelInfo = numbering.getLevel(numId, lvl);
      const formatted = formatOoxmlCounter(
        value,
        useLegalNumbering ? "decimal" : levelInfo?.numFmt || "decimal",
      );
      computedMarker = computedMarker.replaceAll(placeholder, formatted);
    }
  }

  listRendering.marker = computedMarker;
};
