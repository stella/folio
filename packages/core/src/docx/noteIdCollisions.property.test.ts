import { expect, test } from "bun:test";
import fc from "fast-check";
import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";
import { buildPatchedNotePartXml } from "./selectiveXmlPatch";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";

test(
  "duplicate note edits retain the first owner and deletion cannot revive a later owner",
  () => {
    fc.assert(
      fc.property(
        fc.constantFrom("footnote", "endnote"),
        fc.integer({ min: 1, max: 16 }),
        fc.constantFrom("w", "alt"),
        fc.boolean(),
        (kind, repetitions, prefix, nestedDecoy) => {
          type NoteRootOptions = { content: string; spelling?: string };
          const wrap = ({ content, spelling = prefix }: NoteRootOptions) =>
            `<${spelling}:${kind}s xmlns:${spelling}="${W}" xmlns:w14="${W14}">${content}</${spelling}:${kind}s>`;
          type NoteMarkupOptions = { text: string; paraId: string; spelling?: string };
          const note = ({ text, paraId, spelling = prefix }: NoteMarkupOptions) =>
            `<${spelling}:${kind} ${spelling}:id="23"><${spelling}:p w14:paraId="${paraId}"><${spelling}:r><${spelling}:t>${text}</${spelling}:t></${spelling}:r></${spelling}:p></${spelling}:${kind}>`;
          const later = Array.from({ length: repetitions }, (_, index) =>
            note({ text: `Later ${index}`, paraId: (index + 1).toString(16).padStart(8, "0") }),
          );
          const decoy = nestedDecoy
            ? `<ext:${kind}s xmlns:ext="urn:note-extension">${note({ text: "Decoy", paraId: "FFFFFFFF" })}</ext:${kind}s>`
            : "";
          const originalXml = wrap({
            content: decoy + note({ text: "First", paraId: "15AB2300" }) + later.join(""),
          });
          const baselineXml = wrap({
            content: note({ text: "First", paraId: "15AB2300", spelling: "w" }),
            spelling: "w",
          });
          const serializedXml = wrap({
            content: note({ text: "Changed", paraId: "15AB2300", spelling: "w" }),
            spelling: "w",
          });
          const edited = buildPatchedNotePartXml({
            originalXml,
            baselineXml,
            serializedXml,
            replacementXml: serializedXml,
            elementName: kind,
            changedParaIds: new Set(["15AB2300"]),
          });
          expect(edited.type).toBe("patched");
          if (edited.type !== "patched") return;
          expect(edited.xml).toContain(">Changed</");
          expect(edited.xml).not.toContain(">First</");
          for (const retained of later) expect(edited.xml).toContain(retained);
          expect(edited.xml).toContain(decoy);
          const deleted = buildPatchedNotePartXml({
            originalXml,
            baselineXml,
            serializedXml: wrap({ content: "", spelling: "w" }),
            replacementXml: wrap({ content: "", spelling: "w" }),
            elementName: kind,
            changedParaIds: new Set(["15AB2300"]),
          });
          expect(deleted).toEqual({ type: "patched", xml: wrap({ content: decoy }) });
        },
      ),
      propertyConfig(),
    );
  },
  propertyTestTimeout(5_000),
);
