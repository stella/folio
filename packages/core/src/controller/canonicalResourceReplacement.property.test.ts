import { panic } from "better-result";
import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { EditorState } from "prosemirror-state";

import { validateDocxPackage } from "@stll/docx-core";
import { packageResourcesOpOf } from "@stll/docx-core/ops";
import { ZipArchive } from "@stll/docx-core/zip";

import { describePackageDifferences } from "../../../../scripts/lib/corpus-invariants/model-equality";
import { assertExactModel } from "../../../../test/exactModel";
import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { createCanonicalSession } from "../canonical/canonicalSession";
import { serializeCanonicalSave } from "../docx/canonicalSave";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import { createSimpleDocument } from "../docx/serializer/documentSerializer";
import { schema } from "../prosemirror/schema";
import type { CanonicalSaveSnapshot } from "../types/canonicalSave";
import type { Document } from "../types/document";
import {
  FOLIO_DOCX_SERIALIZATION_MODE,
  type FolioDocxSerializationMode,
} from "../types/docxSerialization";

setDefaultTimeout(propertyTestTimeout(60_000));

const MEDIA_PATH = "word/media/canonical-resource.png";
const PNG = Buffer.from(
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+a2ioAAAAASUVORK5CYII=",
  "base64",
);
const STYLE_ID = "Replaceable";
// Markup the model does not hold: a replaced definition must keep it.
const CUSTOM_STYLE = `<w:style w:type="paragraph" w:customStyle="1" w:styleId="${STYLE_ID}"><w:name w:val="${STYLE_ID}"/><w:aliases w:val="Alias"/><w:basedOn w:val="Normal"/><w:rsid w:val="00AB12CD"/><w:pPr><w:keepNext/></w:pPr></w:style>`;
const LATENT_EXCEPTION = `<w:lsdException w:name="Normal" w:qFormat="1"/>`;
const LATENT_STYLES = `<w:latentStyles w:defQFormat="0" w:count="2">${LATENT_EXCEPTION}</w:latentStyles>`;

const ELEMENT = {
  style: new RegExp(
    `<w:style\\b[^>]*w:styleId="${STYLE_ID}"[^>]*>[\\s\\S]*?</w:style>`,
    "u",
  ),
  docDefaults: /<w:docDefaults>[\s\S]*?<\/w:docDefaults>/u,
  latentStyles: /<w:latentStyles\b[^>]*?(?:\/>|>[\s\S]*?<\/w:latentStyles>)/u,
} as const;

const KINDS = [
  "style",
  "styleRemoval",
  "docDefaults",
  "latentStyles",
  "media",
  "mediaType",
  "mediaRemoval",
] as const;
type Kind = (typeof KINDS)[number];

const STYLE_ELEMENT: Record<Kind, keyof typeof ELEMENT | undefined> = {
  style: "style",
  styleRemoval: "style",
  docDefaults: "docDefaults",
  latentStyles: "latentStyles",
  media: undefined,
  mediaType: undefined,
  mediaRemoval: undefined,
};

const textOf = async (zip: ZipArchive, path: string): Promise<string> =>
  (await zip.entry(path)?.text()) ?? panic(`The package is missing ${path}`);

/** The part with the first match of `pattern` cut out, and the match. */
const cut = (xml: string, pattern: RegExp) => {
  const match = pattern.exec(xml);
  return match
    ? {
        rest:
          xml.slice(0, match.index) + xml.slice(match.index + match[0].length),
        element: match[0],
      }
    : { rest: xml, element: undefined };
};

const openSource = async (): Promise<Document> => {
  const zip = await ZipArchive.read(
    await createDocx(createSimpleDocument([{ text: "seed" }])),
  );
  const styles = await textOf(zip, "word/styles.xml");
  const firstStyle = styles.indexOf("<w:style ");
  const close = styles.lastIndexOf("</w:styles>");
  if (firstStyle < 0 || close < 0 || styles.includes("<w:latentStyles"))
    return panic("The generated source has an unexpected style table");
  zip.set(
    "word/styles.xml",
    styles.slice(0, firstStyle) +
      LATENT_STYLES +
      styles.slice(firstStyle, close) +
      CUSTOM_STYLE +
      styles.slice(close),
  );
  zip.set(MEDIA_PATH, PNG);
  const types = await textOf(zip, "[Content_Types].xml");
  if (!types.includes('Extension="png"'))
    zip.set(
      "[Content_Types].xml",
      types.replace(
        "</Types>",
        '<Default Extension="png" ContentType="image/png"/></Types>',
      ),
    );
  return parseDocx((await zip.generate({ compression: "store" })).buffer, {
    preloadFonts: false,
  });
};

/** What a host sets: the package after replacing one resource. */
const replacedOf = (document: Document, kind: Kind, value: number) => {
  const styles =
    document.package.styles ?? panic("The generated source is missing styles");
  const media =
    document.package.media?.get(MEDIA_PATH) ??
    panic("The generated source is missing media");
  const next: Document = { ...document, package: { ...document.package } };
  switch (kind) {
    case "style":
      next.package.styles = {
        ...styles,
        styles: styles.styles.map((style) =>
          style.styleId === STYLE_ID
            ? {
                ...style,
                name: `Replacement ${value}`,
                rPr: { ...style.rPr, bold: value % 2 === 0 },
              }
            : style,
        ),
      };
      break;
    case "styleRemoval":
      next.package.styles = {
        ...styles,
        styles: styles.styles.filter(({ styleId }) => styleId !== STYLE_ID),
      };
      break;
    case "docDefaults":
      next.package.styles = {
        ...styles,
        docDefaults: {
          ...styles.docDefaults,
          pPr: { ...styles.docDefaults?.pPr, keepNext: true },
        },
      };
      break;
    case "latentStyles":
      next.package.styles = {
        ...styles,
        latentStyles: { ...styles.latentStyles, defUIPriority: value },
      };
      break;
    case "media":
    case "mediaType": {
      const signature =
        kind === "media"
          ? [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]
          : [0xff, 0xd8, 0xff, 0xe0];
      next.package.media = new Map(document.package.media);
      next.package.media.set(MEDIA_PATH, {
        path: media.path,
        filename: media.filename,
        mimeType: kind === "media" ? "image/png" : "image/jpeg",
        data: new Uint8Array([...signature, value, value]).buffer,
      });
      break;
    }
    case "mediaRemoval":
      next.package.media = new Map(document.package.media);
      next.package.media.delete(MEDIA_PATH);
      break;
    default: {
      const exhaustive: never = kind;
      return exhaustive;
    }
  }
  return next;
};

/**
 * The parser also keys each media file by its story-relative alias; the
 * package path is the one key that names the part, and the one a host's
 * resource operation can set.
 */
const packageKeyed = (document: Document): Document => ({
  ...document,
  package: {
    ...document.package,
    ...(document.package.media && {
      media: new Map(
        [...document.package.media].filter(([key]) => key.startsWith("word/")),
      ),
    }),
  },
});

const SAVES = Object.values(FOLIO_DOCX_SERIALIZATION_MODE).flatMap(
  (mode: FolioDocxSerializationMode) =>
    [true, false].map((selectiveSave) => ({ mode, selectiveSave })),
);

const saveAll = async (session: {
  captureSaveSnapshot: () => CanonicalSaveSnapshot;
}) =>
  Promise.all(
    SAVES.map(async ({ mode, selectiveSave }) => {
      const saved = await serializeCanonicalSave({
        snapshot: session.captureSaveSnapshot(),
        featureFlags: { selectiveSave },
        options: { mode },
      });
      expect(await validateDocxPackage(new Uint8Array(saved.buffer))).toEqual({
        valid: true,
      });
      return saved;
    }),
  );

test("a canonical save writes replaced styles, style defaults and media as the session holds them, and nothing else", async () => {
  await assertProperty(
    fc.asyncProperty(
      fc.constantFrom(...KINDS),
      fc.integer({ min: 3, max: 255 }),
      async (kind, value) => {
        const source = await openSource();
        const sourceZip = await ZipArchive.read(
          source.originalBuffer ?? panic("The source has no baseline"),
        );
        const sourceStyles = await textOf(sourceZip, "word/styles.xml");
        const session = createCanonicalSession(source).unwrap();
        let state = EditorState.create({ schema, doc: session.projection.doc });
        const baseline = session.captureSaveSnapshot();
        const replaced = replacedOf(session.document, kind, value);
        const commit = session
          .prepareOperations(state, [
            packageResourcesOpOf({ before: session.document, after: replaced }),
          ])
          .unwrap();
        state = state.apply(commit.transaction);
        commit.publish().unwrap();
        const expected = session.captureSaveSnapshot().document;
        const applied = await saveAll(session);

        // Undo restores the opened package; its saves are the control every
        // replaced save is compared against, part by part.
        const undo = session.prepareUndo(state).unwrap();
        state = state.apply(undo.transaction);
        undo.publish().unwrap();
        assertExactModel(
          session.captureSaveSnapshot().document,
          baseline.document,
        );
        const controls = await saveAll(session);
        const redo = session.prepareRedo(state).unwrap();
        state = state.apply(redo.transaction);
        redo.publish().unwrap();
        const redone = await saveAll(session);

        for (const [index, saved] of [...applied, ...redone].entries()) {
          const control =
            controls[index % controls.length] ?? panic("A save has no control");
          const reopened = await parseDocx(saved.buffer, {
            preloadFonts: false,
          });
          // The reopened package is the session's: the replacement, and the
          // relationships, styles and media it left alone.
          expect(
            describePackageDifferences(
              packageKeyed(expected),
              packageKeyed(reopened),
            ),
          ).toEqual({ messages: [], omitted: 0 });
          expect(
            describePackageDifferences(
              packageKeyed(baseline.document),
              packageKeyed(
                await parseDocx(control.buffer, { preloadFonts: false }),
              ),
            ),
          ).toEqual({ messages: [], omitted: 0 });

          const savedZip = await ZipArchive.read(saved.buffer);
          const controlZip = await ZipArchive.read(control.buffer);
          const changed = new Set(["docProps/core.xml"]);
          if (STYLE_ELEMENT[kind] !== undefined) changed.add("word/styles.xml");
          if (kind === "media" || kind === "mediaType") changed.add(MEDIA_PATH);
          if (kind === "mediaType") changed.add("[Content_Types].xml");
          expect(savedZip.paths().toSorted()).toEqual(
            controlZip
              .paths()
              .filter((path) => kind !== "mediaRemoval" || path !== MEDIA_PATH)
              .toSorted(),
          );
          for (const path of savedZip.paths()) {
            if (changed.has(path)) continue;
            expect([path, await savedZip.entry(path)?.bytes()]).toEqual([
              path,
              await controlZip.entry(path)?.bytes(),
            ]);
          }

          const element = STYLE_ELEMENT[kind];
          if (element !== undefined) {
            // Every byte of the style table but the replaced element is the
            // source's.
            const savedStyles = await textOf(savedZip, "word/styles.xml");
            expect(await textOf(controlZip, "word/styles.xml")).toBe(
              sourceStyles,
            );
            const before = cut(sourceStyles, ELEMENT[element]);
            const after = cut(savedStyles, ELEMENT[element]);
            expect(after.rest).toBe(before.rest);
            if (kind === "styleRemoval") expect(after.element).toBeUndefined();
            if (kind === "style") {
              const written =
                after.element ?? panic("The replaced style is missing");
              for (const kept of [
                'w:customStyle="1"',
                '<w:aliases w:val="Alias"/>',
                '<w:rsid w:val="00AB12CD"/>',
              ])
                expect(written).toContain(kept);
              const order = [
                "<w:name ",
                "<w:aliases ",
                "<w:basedOn ",
                "<w:rsid ",
                "<w:pPr>",
                "<w:rPr>",
              ].map((tag) => written.indexOf(tag));
              expect(order).toEqual(order.toSorted((a, b) => a - b));
            }
            if (kind === "latentStyles")
              expect(after.element).toContain(LATENT_EXCEPTION);
          }
          if (kind === "media" || kind === "mediaType") {
            const media =
              expected.package.media?.get(MEDIA_PATH) ??
              panic("The replaced media is missing");
            expect(await savedZip.entry(MEDIA_PATH)?.bytes()).toEqual(
              new Uint8Array(media.data),
            );
          }
          if (kind === "mediaType") {
            const override = `<Override PartName="/${MEDIA_PATH}" ContentType="image/jpeg"/>`;
            const types = await textOf(savedZip, "[Content_Types].xml");
            expect(types).toContain(override);
            expect(types.replace(override, "")).toBe(
              await textOf(controlZip, "[Content_Types].xml"),
            );
          }
        }
      },
    ),
    {
      id: "a canonical save writes replaced styles, style defaults and media as the session holds them, and nothing else",
      numRuns: 21,
    },
  );
});
