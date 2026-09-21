/**
 * The containers routed through the shared child dispatcher.
 *
 * A row is one handler map. Its members are the `w:`-local element names and
 * the complex types the census charges their pairs to, and the generated name
 * list is their union: one walker serves `w:body`, a header, a cell, an SDT's
 * content and a note, so one map has to be total over everything any of them
 * may hold. A union over-approximates for each member alone, which is the safe
 * direction — a decision recorded for a child a given container cannot hold
 * costs a line; a child with no decision costs the markup.
 *
 * A container joins this table when its parser is migrated. Adding a row is
 * how a migration declares itself: the generated union then makes the new
 * handler map's gaps a compile error rather than a silent drop.
 *
 * The table lives here rather than in `scripts/generate-container-children.ts`
 * because that file is a CLI: importing it runs it. Both the generator and the
 * test that binds a sequence row's order to the corpus validator's read this.
 */

export type DispatchedContainer = {
  key: string;
  members: readonly (readonly [element: string, type: string])[];
  /**
   * The children have a declared order, and the generated list is written in
   * it rather than sorted.
   *
   * A property set is the case: `CT_TblPr` and `CT_SectPr` are sequences, and
   * a consumer refuses a `w:tblPr` whose children are in any other order. The
   * order is therefore the generated list itself rather than a second table
   * beside it, so a serializer that sorts by it cannot drift from the set the
   * handler map is total over.
   *
   * `CT_TrPrBase` is the one member whose compositor is a repeated choice, so
   * for `row-properties` the generated list is *a* valid order rather than the
   * only one. Marking it anyway is what gives the row one writer and one sink
   * ordinal; the alternative is a count of modelled siblings, which moves under
   * every capture the day one more property is modelled.
   */
  sequence?: true;
};

export const DISPATCHED_CONTAINERS: readonly DispatchedContainer[] = [
  { key: "w:comment", members: [["comment", "CT_Comment"]] },
  {
    // One walk serves a paragraph, a run-level tracked-change wrapper, a
    // bidirectional wrapper, a smart tag and an inline content control: they
    // share `EG_PContent`/`EG_ContentRunContent` and folio reads them with one
    // function, so one map has to be total over everything any of them holds.
    key: "run-level-content",
    members: [
      ["p", "CT_P"],
      ["ins", "CT_RunTrackChange"],
      ["smartTag", "CT_SmartTagRun"],
      ["customXml", "CT_CustomXmlRun"],
      ["bdo", "CT_BdoContentRun"],
      ["dir", "CT_DirContentRun"],
      ["sdtContent", "CT_SdtContentRun"],
    ],
  },
  {
    // A table and a row-level content control share one walk: folio unwraps
    // the control and splices its rows into the table. A `w:customXml` row
    // wrapper is kept whole rather than unwrapped, so this walk never descends
    // into `CT_CustomXmlRow`; it is a member anyway, because the union is the
    // safe direction and its one extra name — `w:customXmlPr` — then carries a
    // decision instead of falling to a default.
    key: "table-content",
    members: [
      ["tbl", "CT_Tbl"],
      ["sdtContent", "CT_SdtContentRow"],
      ["customXml", "CT_CustomXmlRow"],
    ],
  },
  {
    // A row and a row-level content control share one walk: folio unwraps the
    // control and splices its rows' content into the row, so one map has to be
    // total over everything either may hold.
    key: "row-content",
    members: [
      ["tr", "CT_Row"],
      ["sdtContent", "CT_SdtContentRow"],
    ],
  },
  // A link and a simple field each hold their own subset of `EG_PContent`
  // and each has its own parser, so each gets its own row rather than
  // borrowing `run-level-content`: the union would make a handler map total
  // over names the container cannot hold, and the two parsers would then
  // record decisions for children that never reach them.
  { key: "w:hyperlink", members: [["hyperlink", "CT_Hyperlink"]] },
  { key: "w:fldSimple", members: [["fldSimple", "CT_SimpleField"]] },
  {
    key: "block-content",
    members: [
      ["body", "CT_Body"],
      ["hdr", "CT_HdrFtr"],
      ["ftr", "CT_HdrFtr"],
      ["tc", "CT_Tc"],
      ["sdtContent", "CT_SdtContentBlock"],
      ["footnote", "CT_FtnEdn"],
      ["endnote", "CT_FtnEdn"],
    ],
  },
  {
    // One property set with four owners: a paragraph's own `w:pPr`, the one a
    // style carries, a numbering level's, and the snapshot `w:pPrChange`
    // holds. `CT_PPr` is `CT_PPrBase` plus `w:rPr`, `w:sectPr` and
    // `w:pPrChange`; `CT_PPrGeneral` is `CT_PPrBase` plus `w:pPrChange`, so
    // the merged sequence is the widest of the three and every owner's
    // children carry one decision.
    key: "paragraph-properties",
    members: [
      ["pPr", "CT_PPr"],
      ["pPr", "CT_PPrBase"],
      ["pPr", "CT_PPrGeneral"],
    ],
    sequence: true,
  },
  {
    // A table's own property set, and the one a style carries. `CT_TblPr` is
    // `CT_TblPrBase` plus `w:tblPrChange`, so the two agree on every child
    // they share and the merged sequence is the wider of the two.
    key: "table-properties",
    members: [
      ["tblPr", "CT_TblPr"],
      ["tblPr", "CT_TblPrBase"],
    ],
    sequence: true,
  },
  {
    // A row's table property exceptions, and the snapshot a `w:tblPrExChange`
    // holds. `CT_TblPrEx` is `CT_TblPrExBase` plus that change, the same way
    // `CT_TblPr` is `CT_TblPrBase` plus `w:tblPrChange`. Their nine shared
    // children are the middle of `CT_TblPrBase` — the table properties a row
    // may override — but this is a row of its own rather than two more members
    // of `table-properties`: the exceptions have a parser of their own, and
    // the union would make its handler map total over `w:tblStyle`,
    // `w:tblCaption` and six more names a `w:tblPrEx` cannot hold.
    key: "table-property-exceptions",
    members: [
      ["tblPrEx", "CT_TblPrEx"],
      ["tblPrEx", "CT_TblPrExBase"],
    ],
    sequence: true,
  },
  {
    // A row's own property set, and the snapshot a `w:trPrChange` holds.
    // `CT_TrPr` is `CT_TrPrBase` plus the row's structural revision and that
    // change.
    //
    // `CT_TrPrBase` is a repeated `choice` rather than a sequence, so its
    // twelve children may be written in any order: the generated list is *a*
    // valid order, not the only one, and writing the set in it normalises what
    // folio already normalised by writing its properties in the order of the
    // serializer's statements. `CT_TrPr`'s own three children close a sequence
    // over the base, so `w:ins`, `w:del` and `w:trPrChange` do have to come
    // last, in that order, and the generated list puts them there.
    key: "row-properties",
    members: [
      ["trPr", "CT_TrPr"],
      ["trPr", "CT_TrPrBase"],
    ],
    sequence: true,
  },
  {
    // A cell's own property set, and the snapshot a `w:tcPrChange` holds.
    // `CT_TcPr` is `CT_TcPrInner` plus that change, and `CT_TcPrInner` is
    // `CT_TcPrBase` plus `EG_CellMarkupElements` — the cell's structural
    // revision. Every compositor in the chain is a sequence but that group,
    // which is a choice of three mutually exclusive revisions, so the order is
    // the schema's and a consumer refuses a `w:tcPr` written in another.
    key: "cell-properties",
    members: [
      ["tcPr", "CT_TcPr"],
      ["tcPr", "CT_TcPrInner"],
    ],
    sequence: true,
  },
  {
    // One property set with four owners: a run's own `w:rPr`, the paragraph
    // mark's, and the snapshot each of their `w:rPrChange` records holds.
    // `CT_RPr` is `EG_RPrBase` plus `w:rPrChange`; `CT_ParaRPr` opens with
    // `EG_ParaRPrTrackChanges` on top of that, so the merged sequence is the
    // widest of the four and every owner's children carry one decision.
    key: "run-properties",
    members: [
      ["rPr", "CT_ParaRPr"],
      ["rPr", "CT_RPr"],
      ["rPr", "CT_ParaRPrOriginal"],
      ["rPr", "CT_RPrOriginal"],
    ],
    sequence: true,
  },
  {
    // A section's properties, and the snapshot a `w:sectPrChange` holds.
    // `CT_SectPrBase` is `CT_SectPr` without the two header/footer references
    // that open it and without the change that closes it.
    key: "section-properties",
    members: [
      ["sectPr", "CT_SectPr"],
      ["sectPr", "CT_SectPrBase"],
    ],
    sequence: true,
  },
];
