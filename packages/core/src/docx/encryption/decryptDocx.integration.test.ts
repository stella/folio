/**
 * Integration tests for Agile-encrypted .docx open.
 *
 * Fixture `encrypted-agile-example.docx` is the `example_password.docx` file from the
 * MIT-licensed msoffcrypto-tool test corpus (password: `Password1234_`).
 * @see https://github.com/nolze/msoffcrypto-tool/tree/master/tests/inputs
 */
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import JSZip from "jszip";

import { parseDocx } from "../parser";
import { DOCX_ENCRYPTION_ERROR_CODES, DocxEncryptionError } from "./errors";
import { decryptDocxIfNeeded } from "./decryptDocx";

const __dirname = dirname(fileURLToPath(import.meta.url));
const FIXTURE_PATH = join(__dirname, "fixtures", "encrypted-agile-example.docx");
const CORRECT_PASSWORD = "Password1234_";

const loadFixture = (): Uint8Array => new Uint8Array(readFileSync(FIXTURE_PATH));

const isZip = (data: Uint8Array): boolean =>
  data.length >= 4 &&
  data[0] === 0x50 &&
  data[1] === 0x4b &&
  data[2] === 0x03 &&
  data[3] === 0x04;

describe("decryptDocxIfNeeded — encrypted fixture integration", () => {
  test("decrypts an encrypted .docx with the correct password", async () => {
    const result = await decryptDocxIfNeeded(loadFixture(), { password: CORRECT_PASSWORD });

    expect(result.wasEncrypted).toBe(true);
    expect(isZip(new Uint8Array(result.data))).toBe(true);
  }, 30_000);

  test("throws PASSWORD_REQUIRED when no password is provided", async () => {
    try {
      await decryptDocxIfNeeded(loadFixture());
      expect.unreachable("Expected DocxEncryptionError");
    } catch (error) {
      expect(error).toBeInstanceOf(DocxEncryptionError);
      expect((error as DocxEncryptionError).code).toBe(
        DOCX_ENCRYPTION_ERROR_CODES.PASSWORD_REQUIRED,
      );
    }
  });

  test("throws PASSWORD_INVALID for the wrong password", async () => {
    try {
      await decryptDocxIfNeeded(loadFixture(), { password: "wrong-password" });
      expect.unreachable("Expected DocxEncryptionError");
    } catch (error) {
      expect(error).toBeInstanceOf(DocxEncryptionError);
      expect((error as DocxEncryptionError).code).toBe(
        DOCX_ENCRYPTION_ERROR_CODES.PASSWORD_INVALID,
      );
    }
  }, 30_000);

  test("parseDocx opens the encrypted fixture with password", async () => {
    const doc = await parseDocx(loadFixture(), {
      password: CORRECT_PASSWORD,
      preloadFonts: false,
      detectVariables: false,
    });

    expect(doc.package.document.content.length).toBeGreaterThan(0);
    expect(doc.warnings?.some((warning) => warning.includes("password-protected"))).toBe(true);

    const zip = await JSZip.loadAsync(doc.originalBuffer!);
    const documentXml = await zip.file("word/document.xml")!.async("text");
    expect(documentXml).toContain("Lorem ipsum");
  }, 30_000);
});
