import { createCfbReader } from "./cfbReader";
import { DOCX_ENCRYPTION_ERROR_CODES, DocxEncryptionError } from "./errors";

export type OleEncryptionStreams = {
  encryptionInfo: Uint8Array;
  encryptedPackage: Uint8Array;
};

export const extractEncryptionStreams = (data: ArrayBuffer | Uint8Array): OleEncryptionStreams => {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);

  let reader: ReturnType<typeof createCfbReader>;
  try {
    reader = createCfbReader(bytes);
  } catch (error) {
    throw new DocxEncryptionError({
      code: DOCX_ENCRYPTION_ERROR_CODES.DECRYPTION_FAILED,
      message: "Failed to parse OLE/CFB container",
      cause: error,
    });
  }

  const encryptionInfo = reader.getStream("/EncryptionInfo");
  if (!encryptionInfo) {
    throw new DocxEncryptionError({
      code: DOCX_ENCRYPTION_ERROR_CODES.DECRYPTION_FAILED,
      message: "CFB container missing EncryptionInfo stream",
    });
  }

  const encryptedPackage = reader.getStream("/EncryptedPackage");
  if (!encryptedPackage) {
    throw new DocxEncryptionError({
      code: DOCX_ENCRYPTION_ERROR_CODES.DECRYPTION_FAILED,
      message: "CFB container missing EncryptedPackage stream",
    });
  }

  return { encryptionInfo, encryptedPackage };
};
