/**
 * Synthetic WordprocessingML packages with the parts the package-rewrite
 * tests need: stories, styles, settings, comments and their companion
 * parts, and document properties. Each part is optional and spelled by the
 * test, so a test states exactly what the package holds.
 */

import { writeFile } from "node:fs/promises";
import path from "node:path";

import { ZipArchive } from "@stll/docx-core/zip";

export const W_TRANSITIONAL =
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
export const W_STRICT = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml";
const W15_NS = "http://schemas.microsoft.com/office/word/2012/wordml";
const R_NS =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG_REL_NS =
  "http://schemas.openxmlformats.org/package/2006/relationships";
const REL =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const WML = "application/vnd.openxmlformats-officedocument.wordprocessingml";
const DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

/** Namespace declarations every WordprocessingML part of a fixture carries. */
const wNamespaces = (w: string): string =>
  `xmlns:w="${w}" xmlns:w14="${W14_NS}" xmlns:w15="${W15_NS}" xmlns:r="${R_NS}"`;

export type PackageFixture = {
  /** The children of `w:body`, without the final `w:sectPr`. */
  body: string;
  /** The body's final section properties (default: an empty `w:sectPr`). */
  sectPr?: string;
  /** Header parts by relationship id: each the children of `w:hdr`. */
  headers?: Readonly<Record<string, string>>;
  /** Footer parts by relationship id: each the children of `w:ftr`. */
  footers?: Readonly<Record<string, string>>;
  /** The children of `w:styles`. */
  styles?: string;
  /** The children of `w:settings`. */
  settings?: string;
  /** The children of `w:numbering`. */
  numbering?: string;
  /** The children of `w:comments`. */
  comments?: string;
  /** Whole companion parts of the comments part, written as given. */
  commentsExtended?: string;
  commentsIds?: string;
  commentsExtensible?: string;
  /** The children of `w15:people`. */
  people?: string;
  /** The children of `cp:coreProperties`. */
  core?: string;
  /** The children of the extended `Properties`. */
  app?: string;
  /** The WordprocessingML namespace (default: Transitional). */
  namespace?: string;
};

type Part = {
  name: string;
  contentType: string;
  relationshipType: string;
  xml: string;
  /** Related from the package (`_rels/.rels`) rather than the main document. */
  packageLevel?: boolean;
  id?: string;
};

const wPart = (w: string, root: string, children: string, extra = ""): string =>
  `${DECLARATION}<w:${root} ${wNamespaces(w)}${extra}>${children}</w:${root}>`;

/** Build the package a fixture describes. */
export const buildPackage = async (
  fixture: PackageFixture,
): Promise<Uint8Array> => {
  const w = fixture.namespace ?? W_TRANSITIONAL;
  const parts: Part[] = [];
  for (const [id, children] of Object.entries(fixture.headers ?? {})) {
    parts.push({
      id,
      name: `word/header-${id}.xml`,
      contentType: `${WML}.header+xml`,
      relationshipType: `${REL}/header`,
      xml: wPart(w, "hdr", children),
    });
  }
  for (const [id, children] of Object.entries(fixture.footers ?? {})) {
    parts.push({
      id,
      name: `word/footer-${id}.xml`,
      contentType: `${WML}.footer+xml`,
      relationshipType: `${REL}/footer`,
      xml: wPart(w, "ftr", children),
    });
  }
  const optional: readonly [
    keyof PackageFixture,
    Omit<Part, "xml">,
    (children: string) => string,
  ][] = [
    [
      "styles",
      {
        name: "word/styles.xml",
        contentType: `${WML}.styles+xml`,
        relationshipType: `${REL}/styles`,
      },
      (children) => wPart(w, "styles", children),
    ],
    [
      "settings",
      {
        name: "word/settings.xml",
        contentType: `${WML}.settings+xml`,
        relationshipType: `${REL}/settings`,
      },
      (children) => wPart(w, "settings", children),
    ],
    [
      "numbering",
      {
        name: "word/numbering.xml",
        contentType: `${WML}.numbering+xml`,
        relationshipType: `${REL}/numbering`,
      },
      (children) => wPart(w, "numbering", children),
    ],
    [
      "comments",
      {
        name: "word/comments.xml",
        contentType: `${WML}.comments+xml`,
        relationshipType: `${REL}/comments`,
      },
      (children) => wPart(w, "comments", children),
    ],
    [
      "commentsExtended",
      {
        name: "word/commentsExtended.xml",
        contentType: `${WML}.commentsExtended+xml`,
        relationshipType:
          "http://schemas.microsoft.com/office/2011/relationships/commentsExtended",
      },
      (xml) => xml,
    ],
    [
      "commentsIds",
      {
        name: "word/commentsIds.xml",
        contentType: `${WML}.commentsIds+xml`,
        relationshipType:
          "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds",
      },
      (xml) => xml,
    ],
    [
      "commentsExtensible",
      {
        name: "word/commentsExtensible.xml",
        contentType: `${WML}.commentsExtensible+xml`,
        relationshipType:
          "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible",
      },
      (xml) => xml,
    ],
    [
      "people",
      {
        name: "word/people.xml",
        contentType: `${WML}.people+xml`,
        relationshipType:
          "http://schemas.microsoft.com/office/2011/relationships/people",
      },
      (children) =>
        `${DECLARATION}<w15:people xmlns:w15="${W15_NS}">${children}</w15:people>`,
    ],
    [
      "core",
      {
        name: "docProps/core.xml",
        contentType:
          "application/vnd.openxmlformats-package.core-properties+xml",
        relationshipType:
          "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
        packageLevel: true,
      },
      (children) =>
        `${DECLARATION}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">${children}</cp:coreProperties>`,
    ],
    [
      "app",
      {
        name: "docProps/app.xml",
        contentType:
          "application/vnd.openxmlformats-officedocument.extended-properties+xml",
        relationshipType: `${REL}/extended-properties`,
        packageLevel: true,
      },
      (children) =>
        `${DECLARATION}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">${children}</Properties>`,
    ],
  ];
  for (const [key, part, wrap] of optional) {
    const value = fixture[key];
    if (typeof value === "string") parts.push({ ...part, xml: wrap(value) });
  }

  const zip = new ZipArchive();
  zip.set(
    "[Content_Types].xml",
    `${DECLARATION}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      `<Override PartName="/word/document.xml" ContentType="${WML}.document.main+xml"/>` +
      parts
        .map(
          ({ name, contentType }) =>
            `<Override PartName="/${name}" ContentType="${contentType}"/>`,
        )
        .join("") +
      "</Types>",
  );
  const relationship = (
    { id, relationshipType, name, packageLevel }: Part,
    index: number,
  ): string =>
    `<Relationship Id="${id ?? `rIdPart${index}`}" Type="${relationshipType}" Target="${packageLevel === true ? name : name.slice("word/".length)}"/>`;
  zip.set(
    "_rels/.rels",
    `${DECLARATION}<Relationships xmlns="${PKG_REL_NS}">` +
      `<Relationship Id="rIdMain" Type="${REL}/officeDocument" Target="word/document.xml"/>` +
      parts
        .map((part, index) => ({ part, index }))
        .filter(({ part }) => part.packageLevel === true)
        .map(({ part, index }) => relationship(part, index))
        .join("") +
      "</Relationships>",
  );
  zip.set(
    "word/_rels/document.xml.rels",
    `${DECLARATION}<Relationships xmlns="${PKG_REL_NS}">` +
      parts
        .map((part, index) => ({ part, index }))
        .filter(({ part }) => part.packageLevel !== true)
        .map(({ part, index }) => relationship(part, index))
        .join("") +
      "</Relationships>",
  );
  for (const { name, xml } of parts) zip.set(name, xml);
  zip.set(
    "word/document.xml",
    wPart(
      w,
      "document",
      `<w:body>${fixture.body}${fixture.sectPr ?? "<w:sectPr/>"}</w:body>`,
    ),
  );
  return await zip.generate({ compression: "store" });
};

/** Write a fixture package into `dir` and return its path. */
export const writePackage = async (
  dir: string,
  name: string,
  fixture: PackageFixture,
): Promise<string> => {
  const filePath = path.join(dir, name);
  await writeFile(filePath, await buildPackage(fixture));
  return filePath;
};

/** Every part of a package as text, by name. */
export const packageTexts = async (
  bytes: Uint8Array,
): Promise<ReadonlyMap<string, string>> => {
  const zip = await ZipArchive.read(bytes);
  const texts = new Map<string, string>();
  for (const entry of zip.entries()) texts.set(entry.path, await entry.text());
  return texts;
};
