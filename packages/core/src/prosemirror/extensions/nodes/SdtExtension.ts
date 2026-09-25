/**
 * SDT Extension — inline content control node (Structured Document Tag)
 *
 * Represents OOXML inline SDTs as an inline node wrapping text content.
 * Supports: richText, plainText, date, dropdown, comboBox, checkbox.
 */

import { expectSdtAttrs } from "../../attrs";
import { onOffFromDataset } from "../../conversion/sdtAttrs";
import { createNodeExtension } from "../create";

/** The editor's node for an inline `w:sdt`. */
export const INLINE_CONTENT_CONTROL_NODE_NAME = "sdt";

export const SdtExtension = createNodeExtension({
  name: INLINE_CONTENT_CONTROL_NODE_NAME,
  schemaNodeName: INLINE_CONTENT_CONTROL_NODE_NAME,
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
      /** `w:showingPlcHdr`: `null` when the control states nothing. */
      showingPlaceholder: { default: null },
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
       * The `w:sdtPr` children folio does not model, at their schema ordinal,
       * and the verbatim `<w:sdtEndPr>` the parser captured, so a control keeps
       * what it carries after a save. Mirrors the block-SDT node.
       */
      _preserved: { default: null },
      rawEndPropertiesXml: { default: null },
      /**
       * `w:sdtEndPr` as a record, so a control the editor rebuilds still
       * writes its end mark once the captured bytes are gone.
       */
      endProperties: { default: null },
      /**
       * The revisions whose element encloses the control (`w:ins > w:sdt`),
       * as opposed to revising the content inside it (`w:sdt > w:ins`). Both
       * put the same marks on the leaves; see `contentControlRevisions.ts`.
       */
      _docxEnclosingRevisionIds: { default: null },
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
            showingPlaceholder: onOffFromDataset(el.dataset["showingPlaceholder"]),
            dateFormat: el.dataset["dateFormat"] || null,
            dateValueISO: el.dataset["dateValueIso"] || null,
            listItems: el.dataset["listItems"] || null,
            dropdownLastValue: el.dataset["dropdownLastValue"] || null,
            checked: onOffFromDataset(el.dataset["checked"]),
            // Raw XML is preserved on the model, not the DOM; consumers that
            // round-trip through PM re-attach it from the source.
            _preserved: null,
            rawEndPropertiesXml: null,
            endProperties: null,
            _docxEnclosingRevisionIds: null,
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
      if (attrs.showingPlaceholder !== null && attrs.showingPlaceholder !== undefined) {
        dataAttrs["data-showing-placeholder"] = String(attrs.showingPlaceholder);
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
