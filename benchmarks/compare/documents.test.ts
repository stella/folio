import { expect, test } from "bun:test";

import { buildDocumentPackage, DOCUMENT_CLASSES } from "./documents";
import { zipPackage } from "./package-xml";
import { applyVariant, EDIT_VARIANTS } from "./variants";
import { PACKAGE_VALIDATOR_HINT, resolvePackageValidator } from "./validator";

const validate = resolvePackageValidator();

if (validate === null) {
  test.skip(`generated package schema validation requires ${PACKAGE_VALIDATOR_HINT}`, () => {});
} else {
  test("small generated bases and applicable targets satisfy the OOXML schema", async () => {
    const exercised = new Set<string>();
    for (const documentClass of DOCUMENT_CLASSES) {
      const parts = buildDocumentPackage({ documentClass, size: "s" });
      const baseErrors = validate(await zipPackage(parts));
      expect(baseErrors, `${documentClass}/s/base`).toEqual([]);

      for (const variant of EDIT_VARIANTS) {
        const target = applyVariant({ parts, variant });
        if (target === null) {
          continue;
        }
        exercised.add(variant);
        const errors = validate(await zipPackage(target));
        expect(errors, `${documentClass}/s/${variant}`).toEqual([]);
      }
    }
    expect([...exercised].toSorted()).toEqual([...EDIT_VARIANTS].toSorted());
  }, 120_000);
}
