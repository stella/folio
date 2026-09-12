import { describe, expect, test } from "bun:test";

import type { ParagraphFormatting, StyleDefinitions } from "../types/document";
import { createStyleEngine } from "./styleEngine";
import {
  PARAGRAPH_FORMATTING_PROJECTION_DISPOSITIONS,
  projectTableParagraphPresentation,
  resolveEffectiveParagraphPresentation,
} from "./paragraphPresentation";

const TABLE_OVERLAY_FIELDS = [
  "contextualSpacing",
  "frame",
  "lineSpacing",
  "lineSpacingRule",
  "spaceAfter",
  "spaceBefore",
] as const;

describe("resolved paragraph presentation", () => {
  test("the total disposition map owns the exact supported table-style subset", () => {
    expect(
      Object.entries(PARAGRAPH_FORMATTING_PROJECTION_DISPOSITIONS)
        .filter(([, disposition]) => disposition === "effective-table-overlay")
        .map(([field]) => field)
        .toSorted(),
    ).toEqual(TABLE_OVERLAY_FIELDS);

    const tableFormatting: ParagraphFormatting = {
      alignment: "center",
      bidi: true,
      contextualSpacing: false,
      frame: { width: 720 },
      lineSpacing: 240,
      lineSpacingRule: "auto",
      runProperties: { bold: true },
      spaceAfter: 120,
      spaceBefore: 80,
      suppressLineNumbers: true,
      tabs: [{ position: 360, alignment: "left" }],
    };
    const presentation = projectTableParagraphPresentation(tableFormatting);

    expect(presentation).toEqual({
      overlay: {
        contextualSpacing: false,
        frame: { width: 720 },
        lineSpacing: 240,
        lineSpacingRule: "auto",
        spaceAfter: 120,
        spaceBefore: 80,
      },
      unsupported: [
        { source: "table-style", field: "alignment", value: "center" },
        { source: "table-style", field: "bidi", value: true },
        {
          source: "table-style",
          field: "tabs",
          value: [{ position: 360, alignment: "left" }],
        },
        { source: "table-style", field: "suppressLineNumbers", value: true },
      ],
    });
    expect(presentation?.overlay?.frame).not.toBe(tableFormatting.frame);
    expect(Object.isFrozen(presentation?.overlay)).toBe(true);
    expect(Object.isFrozen(presentation?.overlay?.frame)).toBe(true);
    tableFormatting.frame!.width = 1_440;
    expect(presentation?.overlay?.frame?.width).toBe(720);
  });

  test("resolves defaults, table overlay, style, and direct pPr in order", () => {
    const definitions: StyleDefinitions = {
      docDefaults: {
        pPr: {
          alignment: "left",
          beforeAutospacing: true,
          bidi: true,
          borders: { top: { style: "single", size: 4 } },
          frame: { hAnchor: "margin", width: 480 },
          keepNext: true,
          lineSpacing: 240,
          lineSpacingRule: "auto",
          spaceAfter: 100,
          spaceBefore: 60,
          tabs: [
            { position: 360, alignment: "left" },
            { position: 720, alignment: "center" },
          ],
        },
      },
      styles: [
        {
          styleId: "Clause",
          type: "paragraph",
          pPr: {
            alignment: "center",
            bidi: false,
            frame: { height: 960, width: 600 },
            lineSpacingRule: "atLeast",
            spaceAfter: 180,
            suppressLineNumbers: false,
            tabs: [
              { position: 720, alignment: "right" },
              { position: 1_080, alignment: "center" },
            ],
          },
        },
      ],
    };
    const authored: ParagraphFormatting = {
      alignment: "both",
      afterAutospacing: true,
      beforeAutospacing: false,
      bidi: true,
      borders: { bottom: { style: "double", size: 8 } },
      frame: { vAnchor: "page", width: 900 },
      keepNext: false,
      lineSpacingRule: "exact",
      runProperties: { italic: true },
      spaceBefore: 240,
      spacingExplicit: { before: true },
      styleId: "Clause",
      suppressLineNumbers: true,
      tabs: [
        { position: 360, alignment: "clear" },
        { position: 1_440, alignment: "left" },
      ],
    };
    const tableParagraphPresentation = projectTableParagraphPresentation({
      alignment: "right",
      contextualSpacing: true,
      frame: { hAnchor: "page", vSpace: 120 },
      lineSpacing: 300,
      spaceBefore: 120,
      tabs: [{ position: 2_000, alignment: "right" }],
    });

    const resolved = resolveEffectiveParagraphPresentation({
      authored,
      styleResolver: createStyleEngine(definitions),
      ...(tableParagraphPresentation !== undefined && { tableParagraphPresentation }),
    });

    expect(resolved.inherited).toMatchObject({
      alignment: "center",
      bidi: false,
      contextualSpacing: true,
      frame: { hAnchor: "page", vSpace: 120, height: 960, width: 600 },
      lineSpacing: 300,
      lineSpacingRule: "atLeast",
      spaceAfter: 180,
      spaceBefore: 120,
    });
    expect(resolved.effective).toMatchObject({
      alignment: "both",
      afterAutospacing: true,
      beforeAutospacing: false,
      bidi: true,
      borders: { bottom: { style: "double", size: 8 } },
      contextualSpacing: true,
      frame: {
        hAnchor: "page",
        vAnchor: "page",
        vSpace: 120,
        height: 960,
        width: 900,
      },
      keepNext: false,
      lineSpacing: 300,
      lineSpacingRule: "exact",
      spaceAfter: 180,
      spaceBefore: 240,
      tabs: [
        { position: 360, alignment: "clear" },
        { position: 720, alignment: "right" },
        { position: 1_080, alignment: "center" },
        { position: 1_440, alignment: "left" },
      ],
    });
    expect(resolved.effective.borders).toEqual({
      bottom: { style: "double", size: 8 },
    });
    expect("styleId" in resolved.effective).toBe(false);
    expect("spacingExplicit" in resolved.effective).toBe(false);
    expect("runProperties" in resolved.effective).toBe(false);
    expect("suppressLineNumbers" in resolved.effective).toBe(false);
    expect(resolved.unsupported).toEqual([
      { source: "table-style", field: "alignment", value: "right" },
      {
        source: "table-style",
        field: "tabs",
        value: [{ position: 2_000, alignment: "right" }],
      },
      { source: "inherited", field: "suppressLineNumbers", value: false },
      { source: "direct", field: "suppressLineNumbers", value: true },
    ]);
    expect(authored).toEqual({
      alignment: "both",
      afterAutospacing: true,
      beforeAutospacing: false,
      bidi: true,
      borders: { bottom: { style: "double", size: 8 } },
      frame: { vAnchor: "page", width: 900 },
      keepNext: false,
      lineSpacingRule: "exact",
      runProperties: { italic: true },
      spaceBefore: 240,
      spacingExplicit: { before: true },
      styleId: "Clause",
      suppressLineNumbers: true,
      tabs: [
        { position: 360, alignment: "clear" },
        { position: 1_440, alignment: "left" },
      ],
    });
  });

  test("owns and recursively freezes nested effective values", () => {
    const authored: ParagraphFormatting = {
      borders: { top: { style: "single", color: { rgb: "FF0000" } } },
      frame: { width: 720 },
      numPr: { numId: 4, ilvl: 2 },
      shading: { fill: { themeColor: "accent1" } },
      tabs: [{ position: 360, alignment: "left" }],
    };

    const effective = resolveEffectiveParagraphPresentation({
      authored,
      styleResolver: null,
    }).effective;

    expect(effective).not.toBe(authored);
    expect(effective.borders).not.toBe(authored.borders);
    expect(effective.borders?.top).not.toBe(authored.borders?.top);
    expect(effective.borders?.top?.color).not.toBe(authored.borders?.top?.color);
    expect(effective.frame).not.toBe(authored.frame);
    expect(effective.numPr).not.toBe(authored.numPr);
    expect(effective.shading).not.toBe(authored.shading);
    expect(effective.shading?.fill).not.toBe(authored.shading?.fill);
    expect(effective.tabs).not.toBe(authored.tabs);
    expect(effective.tabs?.at(0)).not.toBe(authored.tabs?.at(0));
    expect(Object.isFrozen(effective)).toBe(true);
    expect(Object.isFrozen(effective.borders)).toBe(true);
    expect(Object.isFrozen(effective.borders?.top)).toBe(true);
    expect(Object.isFrozen(effective.borders?.top?.color)).toBe(true);
    expect(Object.isFrozen(effective.frame)).toBe(true);
    expect(Object.isFrozen(effective.numPr)).toBe(true);
    expect(Object.isFrozen(effective.shading)).toBe(true);
    expect(Object.isFrozen(effective.shading?.fill)).toBe(true);
    expect(Object.isFrozen(effective.tabs)).toBe(true);
    expect(Object.isFrozen(effective.tabs?.at(0))).toBe(true);

    authored.borders!.top!.color!.rgb = "00FF00";
    authored.frame!.width = 1_440;
    authored.numPr!.numId = 9;
    authored.shading!.fill!.themeColor = "accent2";
    authored.tabs!.at(0)!.position = 720;
    expect(effective).toMatchObject({
      borders: { top: { color: { rgb: "FF0000" } } },
      frame: { width: 720 },
      numPr: { numId: 4 },
      shading: { fill: { themeColor: "accent1" } },
      tabs: [{ position: 360 }],
    });
  });

  test("a direct numId owns numbering and numId zero drops style marker indents", () => {
    const styleResolver = createStyleEngine({
      styles: [
        {
          styleId: "Numbered",
          type: "paragraph",
          pPr: {
            indentFirstLine: -360,
            hangingIndent: true,
            indentLeft: 720,
            indentRight: 90,
            numPr: { numId: 12, ilvl: 2 },
          },
        },
      ],
    });

    const removed = resolveEffectiveParagraphPresentation({
      authored: {
        styleId: "Numbered",
        numPr: { numId: 0, ilvl: 0 },
        indentLeft: 357,
      },
      styleResolver,
    }).effective;
    expect(removed.numPr).toEqual({ numId: 0, ilvl: 0 });
    expect(removed.indentLeft).toBe(357);
    expect(removed.indentRight).toBe(90);
    expect(removed.indentFirstLine).toBeUndefined();
    expect(removed.hangingIndent).toBeUndefined();

    const changedIdentity = resolveEffectiveParagraphPresentation({
      authored: { styleId: "Numbered", numPr: { numId: 9 } },
      styleResolver,
    }).effective;
    expect(changedIdentity.numPr).toEqual({ numId: 9 });
    expect(changedIdentity.indentFirstLine).toBe(-360);
    expect(changedIdentity.hangingIndent).toBe(true);

    const changedLevel = resolveEffectiveParagraphPresentation({
      authored: { styleId: "Numbered", numPr: { ilvl: 4 } },
      styleResolver,
    }).effective;
    expect(changedLevel.numPr).toEqual({ numId: 12, ilvl: 4 });
  });
});
