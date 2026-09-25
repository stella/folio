/**
 * Font aliasing for comparison runs. When Chromium cannot resolve a family a
 * document requests, it falls back through Folio's CSS stack while the
 * reference renderer may pick a different face (its own substitute, or a copy
 * of the family bundled with the reference application). Registering that
 * same local face under the requested family gives both renderers identical
 * glyph metrics, so the run measures layout rather than font availability.
 */

import type { FontPair } from "./features";
import type { AliasedFontDefinition } from "./folioExtract";
import { findFamilyFacesForPostscriptName, type FontFaceRecord } from "./fontFaces";
import { fontFamiliesMatch } from "./fontNames";
import type { FontAlias } from "./types";

/** Share of a requested family's paired lines one reference face family must
 * cover before it is treated as the reference's consistent choice. */
export const MIN_FONT_ALIAS_SHARE = 0.8;

export type FontAliasPlan = {
  aliases: FontAlias[];
  fonts: AliasedFontDefinition[];
};

const faceKey = ({ weight, style }: FontFaceRecord): string => `${weight}:${style}`;

/** Plan aliases from text-paired lines of a first comparison run. Only
 * families Chromium could not resolve are aliased, and only to a face family
 * available as a local file. */
export const planFontAliases = (
  pairs: readonly FontPair[],
  faces: readonly FontFaceRecord[],
): FontAliasPlan => {
  const pairsByRequested = new Map<string, { requested: string; pairs: FontPair[] }>();
  for (const pair of pairs) {
    const requested = pair.folioRequestedFont;
    if (requested === undefined || fontFamiliesMatch(requested, pair.folioFont)) continue;
    const key = requested.toLowerCase();
    const group = pairsByRequested.get(key);
    if (group) {
      group.pairs.push(pair);
    } else {
      pairsByRequested.set(key, { requested, pairs: [pair] });
    }
  }

  const aliases: FontAlias[] = [];
  const fonts: AliasedFontDefinition[] = [];
  for (const { requested, pairs: requestedPairs } of pairsByRequested.values()) {
    if (
      requestedPairs.every(({ referenceFont, folioFont }) =>
        fontFamiliesMatch(referenceFont, folioFont),
      )
    ) {
      continue;
    }
    const facesByFamily = new Map<string, { count: number; faces: FontFaceRecord[] }>();
    for (const { referenceFont } of requestedPairs) {
      const familyFaces = findFamilyFacesForPostscriptName(faces, referenceFont);
      const family = familyFaces[0]?.family.toLowerCase();
      if (family === undefined) continue;
      const entry = facesByFamily.get(family);
      if (entry) {
        entry.count += 1;
      } else {
        facesByFamily.set(family, { count: 1, faces: familyFaces });
      }
    }
    const dominant = [...facesByFamily.values()].toSorted((a, b) => b.count - a.count).at(0);
    if (dominant === undefined || dominant.count / requestedPairs.length < MIN_FONT_ALIAS_SHARE) {
      continue;
    }

    const faceFamily = dominant.faces[0]?.family;
    if (faceFamily === undefined) continue;
    const seen = new Set<string>();
    for (const face of dominant.faces) {
      if (seen.has(faceKey(face))) continue;
      seen.add(faceKey(face));
      fonts.push({
        family: requested,
        filePath: face.filePath,
        weight: face.weight,
        ...(face.style === "italic" ? { style: "italic" } : {}),
        reportedFamily: faceFamily,
      });
    }
    aliases.push({ requestedFamily: requested, faceFamily });
  }
  return { aliases, fonts };
};
