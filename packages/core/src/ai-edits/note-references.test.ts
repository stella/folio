import { expect, test } from "bun:test";

import { segmentsAroundNoteReferences } from "./note-references";

test("no split around a preserved reference can spell a new reference", () => {
  for (const marker of ["[^1]", "[^e2]"]) {
    for (let split = 1; split < marker.length; split++) {
      const replacement = marker.slice(0, split) + marker + marker.slice(split);
      expect(
        segmentsAroundNoteReferences(marker, replacement, [{ offset: 0, length: marker.length }]),
      ).toBeNull();
    }
  }
});
