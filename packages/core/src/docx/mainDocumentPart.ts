/**
 * A package names its main document through the `officeDocument` relationship,
 * and nothing requires that name to be `word/document.xml`: some producers write
 * `word/document2.xml`. Every reader and writer here works at the standard path,
 * so a package opens with its main document moved there, and every save reads
 * the package it writes over through {@link readPackageForSave}, which moves it
 * the same way. Which part a package names is answered by
 * {@link mainDocumentPartPath} alone.
 */

import { ZipArchive } from "@stll/docx-core/zip";

import { resolvePackageRelationshipTarget } from "./packageParts";

const STANDARD_PATH = "word/document.xml";
const STANDARD_RELATIONSHIPS_PATH = "word/_rels/document.xml.rels";
const PACKAGE_RELATIONSHIPS_PATH = "_rels/.rels";
const CONTENT_TYPES_PATH = "[content_types].xml";
const OFFICE_DOCUMENT_TYPE = /\/officeDocument["']/u;
const WORD_MAIN_CONTENT_TYPE = /wordprocessingml\.[a-z.]*main\+xml/iu;

const entryNamed = (zip: ZipArchive, lowerPath: string) =>
  zip.paths().find((path) => path.toLowerCase() === lowerPath);

const relationshipsPathOf = (part: string): string => {
  const slash = part.lastIndexOf("/");
  return `${part.slice(0, slash + 1)}_rels/${part.slice(slash + 1)}.rels`;
};

/** A relationship's relative target as an absolute part name. */
const absolutePartPath = (directory: string, relative: string): string => {
  const segments: string[] = [];
  for (const segment of `${directory}${relative}`.split("/")) {
    if (segment === "..") segments.pop();
    else if (segment !== "" && segment !== ".") segments.push(segment);
  }
  return `/${segments.join("/")}`;
};

const escapeForRegExp = (text: string): string =>
  text.replaceAll(/[.*+?^${}()|[\]\\]/gu, String.raw`\$&`);

const targetOf = (element: string): string | null =>
  /\bTarget\s*=\s*["']([^"']+)["']/u.exec(element)?.[1] ?? null;

type DeclaredMainDocument = {
  packagePath: string;
  packageXml: string;
  relationship: string;
  /** The entry the relationship names, spelled as the archive spells it. */
  part: string;
};

const declaredMainDocument = async (
  zip: ZipArchive,
): Promise<DeclaredMainDocument | undefined> => {
  const packagePath = entryNamed(zip, PACKAGE_RELATIONSHIPS_PATH);
  if (packagePath === undefined) return undefined;
  const packageXml = await zip.entry(packagePath)?.text();
  if (packageXml === undefined) return undefined;
  const relationship = [
    ...packageXml.matchAll(/<(?:\w+:)?Relationship\b[^>]*>/gu),
  ]
    .map(([element]) => element)
    .find((element) => OFFICE_DOCUMENT_TYPE.test(element));
  const target = relationship && targetOf(relationship);
  if (!relationship || !target) return undefined;
  const lowerPath = resolvePackageRelationshipTarget(
    target,
    PACKAGE_RELATIONSHIPS_PATH,
  );
  if (lowerPath === undefined) return undefined;
  const part = entryNamed(zip, lowerPath);
  if (part === undefined) return undefined;
  return { packagePath, packageXml, relationship, part };
};

/**
 * The part the package's `officeDocument` relationship names, as the archive
 * spells it, or `undefined` when the package names no part it holds.
 */
export const mainDocumentPartPath = async (
  zip: ZipArchive,
): Promise<string | undefined> => (await declaredMainDocument(zip))?.part;

/**
 * Moves the main document to the standard path. Returns whether it moved.
 * What is not a word processing main part, a theme package say, stays put.
 */
export const moveMainDocumentToStandardPath = async (
  zip: ZipArchive,
): Promise<boolean> => {
  const contentTypesPath = entryNamed(zip, CONTENT_TYPES_PATH);
  if (contentTypesPath === undefined) return false;
  const declaredMain = await declaredMainDocument(zip);
  if (
    declaredMain === undefined ||
    declaredMain.part.toLowerCase() === STANDARD_PATH
  )
    return false;
  const contentTypes = await zip.entry(contentTypesPath)?.text();
  if (contentTypes === undefined) return false;
  const { packagePath, packageXml, relationship, part } = declaredMain;

  const override = new RegExp(
    String.raw`<(?:\w+:)?Override\b[^>]*PartName\s*=\s*["']/${escapeForRegExp(part)}["'][^>]*>`,
    "iu",
  );
  const declared = override.exec(contentTypes)?.[0];
  if (declared === undefined || !WORD_MAIN_CONTENT_TYPE.test(declared)) {
    return false;
  }

  const bytes = await zip.entry(part)?.bytes();
  if (bytes === undefined) return false;

  // What the part's relationships name is relative to where it was.
  const relationshipsPath = entryNamed(
    zip,
    relationshipsPathOf(part).toLowerCase(),
  );
  const relationships =
    relationshipsPath === undefined
      ? undefined
      : await zip.entry(relationshipsPath)?.text();
  const directory = part.slice(0, part.lastIndexOf("/") + 1);
  const anchored =
    relationships === undefined || directory === "word/"
      ? relationships
      : relationships.replaceAll(
          /<(?:\w+:)?Relationship\b[^>]*>/gu,
          (element) =>
            /TargetMode\s*=\s*["']External["']/iu.test(element)
              ? element
              : element.replace(
                  /\bTarget\s*=\s*(["'])(?!\/)([^"']+)\1/u,
                  (_match, quote: string, relative: string) => {
                    const absolute = absolutePartPath(directory, relative);
                    return `Target=${quote}${absolute}${quote}`;
                  },
                ),
        );

  zip.delete(part);
  if (relationshipsPath !== undefined) zip.delete(relationshipsPath);
  const replaced = entryNamed(zip, STANDARD_PATH);
  if (replaced !== undefined) zip.delete(replaced);
  const replacedRelationships = entryNamed(zip, STANDARD_RELATIONSHIPS_PATH);
  if (replacedRelationships !== undefined) zip.delete(replacedRelationships);

  zip.set(STANDARD_PATH, bytes);
  if (anchored !== undefined) zip.set(STANDARD_RELATIONSHIPS_PATH, anchored);
  zip.set(
    packagePath,
    packageXml.replace(
      relationship,
      relationship.replace(
        /\bTarget\s*=\s*(["'])[^"']+\1/u,
        `Target="${STANDARD_PATH}"`,
      ),
    ),
  );
  const standardOverride = new RegExp(
    String.raw`<(?:\w+:)?Override\b[^>]*PartName\s*=\s*["']/${escapeForRegExp(STANDARD_PATH)}["'][^>]*>`,
    "iu",
  );
  zip.set(
    contentTypesPath,
    contentTypes
      .replace(standardOverride, "")
      .replace(
        override,
        declared.replace(
          /PartName\s*=\s*(["'])[^"']+\1/u,
          `PartName="/${STANDARD_PATH}"`,
        ),
      ),
  );
  return true;
};

/**
 * The package a save writes over, with its main document at the standard
 * path. Every save entry that takes the bytes of an opened package reads them
 * here, so bytes handed over as the producer wrote them save to one main part
 * exactly as the bytes the reader kept do.
 */
export const readPackageForSave = async (
  bytes: Uint8Array | ArrayBuffer,
): Promise<ZipArchive> => {
  const zip = await ZipArchive.read(bytes);
  await moveMainDocumentToStandardPath(zip);
  return zip;
};
