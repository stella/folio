// A field-totality gate in a test file is read by no compiler, so the authored
// values for every parser-derived image fact live here, where the package
// typecheck compiles them.
import type { imageDerivedFactsOf } from "../../docx/imageParser";
import type { Image } from "../../types/document";

/** An authored value for every fact the image parser derives. */
export const AUTHORED_IMAGE_FACTS = {
  id: "12345",
  filename: "authored.png",
  mimeType: "image/png",
  pictureNames: {
    name: "Authored picture",
    alt: "Alternative",
    title: "Authored title",
  },
} satisfies Required<Pick<Image, keyof ReturnType<typeof imageDerivedFactsOf>>>;
