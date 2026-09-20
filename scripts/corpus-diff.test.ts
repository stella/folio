import { describe, expect, test } from "bun:test";

import { SIGNATURE_DELTA_KINDS, diffFileSignatures, renderFamilyDeltas } from "./lib/corpus-diff";
import type { CorpusFileSignatures } from "./lib/corpus-file-signatures";
import { fileSignatureRows } from "./lib/corpus-file-signatures";

const ROUND_TRIP = "editor-round-trip | a field went missing | -";
const RESERIALIZE = "reserialize | a field went missing | -";

const row = (id: string, family: string, ...signatures: string[]) =>
  ({ id, family, signatures }) as CorpusFileSignatures;

/**
 * Two runs over four files. Between them one signature arrives, one goes, one
 * spreads to a file it had not reached and loses one it had.
 */
const BEFORE = [
  row("poi/one.docx", "editor-round-trip", ROUND_TRIP),
  row("poi/two.docx", "editor-round-trip", ROUND_TRIP),
  row("poi/two.docx", "reserialize", RESERIALIZE),
];
const AFTER = [
  row("poi/one.docx", "editor-round-trip", ROUND_TRIP),
  row("poi/three.docx", "editor-round-trip", ROUND_TRIP),
  row("poi/four.docx", "editor-round-trip", "editor-round-trip | a new loss | -"),
];

describe("a differential names the files, not an example", () => {
  test("a grown signature reports what reached it and what left", () => {
    const [roundTrip] = diffFileSignatures(BEFORE, AFTER);
    expect(roundTrip?.family).toBe("editor-round-trip");
    expect(roundTrip?.deltas).toContainEqual({
      kind: SIGNATURE_DELTA_KINDS.reachedNow,
      signature: ROUND_TRIP,
      before: 2,
      after: 2,
      reached: ["poi/three.docx"],
      left: ["poi/two.docx"],
    });
  });

  test("a signature only the later run has is introduced", () => {
    const introduced = diffFileSignatures(BEFORE, AFTER)
      .flatMap(({ deltas }) => deltas)
      .filter(({ kind }) => kind === SIGNATURE_DELTA_KINDS.introduced);
    expect(introduced).toEqual([
      {
        kind: SIGNATURE_DELTA_KINDS.introduced,
        signature: "editor-round-trip | a new loss | -",
        before: 0,
        after: 1,
        reached: ["poi/four.docx"],
        left: [],
      },
    ]);
  });

  test("a signature only the earlier run has is fixed, in its own family", () => {
    const reserialize = diffFileSignatures(BEFORE, AFTER).find(
      ({ family }) => family === "reserialize",
    );
    expect(reserialize?.deltas).toEqual([
      {
        kind: SIGNATURE_DELTA_KINDS.fixed,
        signature: RESERIALIZE,
        before: 1,
        after: 0,
        reached: [],
        left: ["poi/two.docx"],
      },
    ]);
  });

  test("two runs that agree report nothing", () => {
    expect(diffFileSignatures(BEFORE, BEFORE)).toEqual([]);
    expect(renderFamilyDeltas([])).toBe("No signature changed its file set.");
  });

  test("the report names every family and every moved file", () => {
    const rendered = renderFamilyDeltas(diffFileSignatures(BEFORE, AFTER));
    expect(rendered).toContain("editor-round-trip: 2 signature(s) moved");
    expect(rendered).toContain("reserialize: 1 signature(s) moved");
    expect(rendered).toContain("        poi/three.docx");
  });
});

describe("a file's rows are the census's own verdict", () => {
  test("failures are grouped by the family that counts them", () => {
    expect(
      fileSignatureRows({ sourceId: "poi", path: "a.docx", sha256: "0".repeat(64) }, [
        { invariant: "editor-round-trip", message: "b", frame: "-" },
        { invariant: "reserialize", message: "a", frame: "-" },
        { invariant: "editor-round-trip", message: "a", frame: "-" },
      ]),
    ).toEqual([
      {
        id: "poi/a.docx",
        family: "editor-round-trip",
        signatures: ["editor-round-trip | a | -", "editor-round-trip | b | -"],
      },
      { id: "poi/a.docx", family: "reserialize", signatures: ["reserialize | a | -"] },
    ]);
  });

  test("a file with no findings contributes no row", () => {
    expect(
      fileSignatureRows({ sourceId: "poi", path: "a.docx", sha256: "0".repeat(64) }, []),
    ).toEqual([]);
  });
});
