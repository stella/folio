/**
 * SDT Extension — inline content control node (Structured Document Tag)
 *
 * Represents OOXML inline SDTs as an inline node wrapping text content.
 * Supports: richText, plainText, date, dropdown, comboBox, checkbox.
 */

import { expectSdtAttrs } from "../../attrs";
import { createNodeExtension } from "../create";

export const SdtExtension = createNodeExtension({
  name: "sdt",
  schemaNodeName: "sdt",
  nodeSpec: {
    inline: true,
    group: "inline",
    content: "inline*",
    attrs: {
      /** SDT type: richText, plainText, date, dropdown, comboBox, checkbox, etc. */
      sdtType: { default: "richText" },
      /** Alias (friendly name) */
      alias: { default: null },
      /** Tag (developer identifier) */
      tag: { default: null },
      /** Numeric `w:id/@w:val`. */
      id: { default: null },
      /** Lock setting */
      lock: { default: null },
      /** Placeholder text */
      placeholder: { default: null },
      /** Whether showing placeholder */
      showingPlaceholder: { default: false },
      /** Date format for date controls */
      dateFormat: { default: null },
      /** ISO 8601 bound date value (`w:date@w:fullDate`). */
      dateValueISO: { default: null },
      /** Dropdown/combobox list items as JSON string */
      listItems: { default: null },
      /** Selected dropdown / comboBox value (`w:dropDownList@w:lastValue`). */
      dropdownLastValue: { default: null },
      /** Checkbox checked state */
      checked: { default: null },
      /**
       * Verbatim `<w:sdtPr>` / `<w:sdtEndPr>` captured by the parser so
       * unmodeled OOXML features (`w:dataBinding`, `w15:*`, custom XML
       * mappings) round-trip after a save, mirroring the block-SDT node.
       */
      rawPropertiesXml: { default: null },
      rawEndPropertiesXml: { default: null },
    },
    parseDOM: [
      {
        tag: "span.docx-sdt",
        getAttrs(dom) {
          if (!(dom instanceof HTMLElement)) {
            return false;
          }
          const el = dom;
          const idRaw = el.dataset["sdtId"];
          const id = idRaw ? Number.parseInt(idRaw, 10) : null;
          return {
            sdtType: el.dataset["sdtType"] || "richText",
            alias: el.dataset["alias"] || null,
            tag: el.dataset["tag"] || null,
            id: id !== null && !Number.isNaN(id) ? id : null,
            lock: el.dataset["lock"] || null,
            placeholder: el.dataset["placeholder"] || null,
            showingPlaceholder: el.dataset["showingPlaceholder"] === "true",
            dateFormat: el.dataset["dateFormat"] || null,
            dateValueISO: el.dataset["dateValueIso"] || null,
            listItems: el.dataset["listItems"] || null,
            dropdownLastValue: el.dataset["dropdownLastValue"] || null,
            checked: (() => {
              if (el.dataset["checked"] === "true") {
                return true;
              }
              if (el.dataset["checked"] === "false") {
                return false;
              }
              return null;
            })(),
            // Raw XML is preserved on the model, not the DOM; consumers that
            // round-trip through PM re-attach it from the source.
            rawPropertiesXml: null,
            rawEndPropertiesXml: null,
          };
        },
      },
    ],
    toDOM(node) {
      const attrs = expectSdtAttrs(node);
      const dataAttrs: Record<string, string> = {
        class: `docx-sdt docx-sdt-${attrs.sdtType}`,
        "data-sdt-type": attrs.sdtType,
      };

      if (attrs.alias) {
        dataAttrs["data-alias"] = attrs.alias;
      }
      if (attrs.tag) {
        dataAttrs["data-tag"] = attrs.tag;
      }
      if (typeof attrs.id === "number") {
        dataAttrs["data-sdt-id"] = String(attrs.id);
      }
      if (attrs.lock) {
        dataAttrs["data-lock"] = attrs.lock;
      }
      if (attrs.placeholder) {
        dataAttrs["data-placeholder"] = attrs.placeholder;
      }
      if (attrs.showingPlaceholder) {
        dataAttrs["data-showing-placeholder"] = "true";
      }
      if (attrs.dateFormat) {
        dataAttrs["data-date-format"] = attrs.dateFormat;
      }
      if (attrs.dateValueISO) {
        dataAttrs["data-date-value-iso"] = attrs.dateValueISO;
      }
      if (attrs.listItems) {
        dataAttrs["data-list-items"] = attrs.listItems;
      }
      if (attrs.dropdownLastValue) {
        dataAttrs["data-dropdown-last-value"] = attrs.dropdownLastValue;
      }
      if (attrs.checked !== undefined) {
        dataAttrs["data-checked"] = String(attrs.checked);
      }

      // Checkbox renders with a checkbox-like indicator
      if (attrs.sdtType === "checkbox") {
        dataAttrs["style"] =
          "border: 1px solid #ccc; border-radius: 3px; padding: 0 2px; display: inline;";
      } else {
        dataAttrs["style"] = "border-bottom: 1px dashed #999; padding: 0 1px; display: inline;";
      }

      return ["span", dataAttrs, 0];
    },
  },
});
