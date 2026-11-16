import type { FolioDocumentOperationResult } from "@stll/folio-core/server";

import type { Mode } from "./operations.ts";

type ParagraphIdentityResults = {
  applied: readonly FolioDocumentOperationResult["applied"][number][];
  mode: Mode;
};

/** Follow only the paragraph retirements the public operation receipt reports. */
export const createParagraphIdentityRemapping = () => {
  const joined = new Map<string, string>();
  return {
    record({ applied, mode }: ParagraphIdentityResults) {
      if (mode !== "direct") return;
      for (const { paragraphIdentity } of applied) {
        if (paragraphIdentity === undefined) continue;
        switch (paragraphIdentity.type) {
          case "joined":
            joined.set(
              paragraphIdentity.retiredBlockId,
              paragraphIdentity.blockId,
            );
            break;
          case "retained":
            break;
          default: {
            const unreachable: never = paragraphIdentity;
            return unreachable;
          }
        }
      }
    },
    resolve(id: string, liveIds: ReadonlySet<string>): string {
      const seen = new Set<string>();
      while (!liveIds.has(id)) {
        if (seen.has(id))
          throw new Error("Paragraph receipts contain a retirement cycle.");
        seen.add(id);
        const survivor = joined.get(id);
        if (survivor === undefined) return id;
        id = survivor;
      }
      return id;
    },
  };
};

export const mapParagraphBlockIds = (
  value: unknown,
  resolve: (id: string) => string,
): unknown => {
  if (Array.isArray(value))
    return value.map((item) => mapParagraphBlockIds(item, resolve));
  if (typeof value !== "object" || value === null) return value;
  return Object.fromEntries(
    Object.entries(value).map(([key, item]) => [
      key,
      key === "blockId" && typeof item === "string"
        ? resolve(item)
        : mapParagraphBlockIds(item, resolve),
    ]),
  );
};
