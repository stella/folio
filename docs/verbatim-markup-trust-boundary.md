# Verbatim markup and the constructed-document writer

folio keeps markup it does not model as bytes, and writes those bytes back into
the package unchanged. On the parse-and-repack path that is sound by
construction: the bytes came out of a package folio just read and bounded, and
`captureVerbatimXml` produced them from a parsed element rather than from a
string somebody supplied.

`serializeDocumentToDocx` (`packages/docx-core/src/serialize/docx.ts`) is a
different path. It writes a package from scratch for a `Document` that was
**built** rather than parsed — today the legal-source compiler's — and its
input is an ordinary TypeScript object literal. Every verbatim slot in that
object is a string the caller chose, and the writer concatenates it into
`word/document.xml`. Nothing checks that the string is one well-formed element,
that it closes what it opens, or that it is bounded.

This is a note, not a change. Nothing here is implemented.

## What is emitted verbatim, and from where

The raw-XML members of the model, by owner:

| Member                                                   | Owner                                                  |
| -------------------------------------------------------- | ------------------------------------------------------ |
| `PreservedXmlContent.xml`                                | `RunContent` — a run child folio does not model        |
| `PreservedInline.xml`                                    | `ParagraphContent` — an inline child it does not model |
| `PreservedBlock.xml`                                     | `BlockContent` — a block child it does not model       |
| `PreservedMarkup.children[].xml`                         | the ordered sink (`Comment.preserved`)                 |
| `MathEquation.ommlXml`                                   | an equation, inline or display                         |
| `DrawingContent.rawXml`                                  | a drawing folio replays rather than rebuilds           |
| `ShapeContent.rawXml`, `Shape.rawXml`                    | a VML or DrawingML shape                               |
| `SdtProperties.rawPropertiesXml` / `rawEndPropertiesXml` | a content control's properties                         |
| `HeaderFooter.rawWatermarkXml`                           | a watermark                                            |
| `TableFormatting.gridChangeXml`                          | `w:tblGridChange`                                      |
| `ParagraphFormatting.numberingChangeXml`                 | `w:numberingChange`                                    |
| `*.sourceXml` (borders, shading, tab stops)              | a property folio re-spells rather than rebuilds        |
| `ImageFrameLocks`-adjacent `docPrExtensions[]`           | `a:ext` children of a `wp:docPr`                       |
| `TextBox.verbatimXml`                                    | a text box folio replays                               |

The constructed path reaches only some of them, because most of its content
switches return `""`. The emitters that do write a caller's string into the
part are:

- `serializeBlockContent` → `case "preservedBlock": return block.xml`
- `serializeRunContent` → `case "preservedXml": return content.xml`
- `serializeParagraphContent` → every other branch returns `""` today,
  including `preservedInline`, `bidiWrapper` and `mathEquation`
- `serializeRunProperties` → the comment at the `style`/`color` branches notes
  the values are typed `string` and may carry preserved unknown OOXML

Two of those three are live today. `preservedInline` is not, and that is the
point at which the surface stops growing by accident: a new capture member
reaching this writer should have to pass the boundary check below before it
gets a branch that emits.

## The one check to add, and where

One validation at the constructed-document entry, not one per emitter. Per-slot
checks are the shape that drifts: a new slot lands with no check, and the check
that exists is in the wrong place to bound the whole document anyway.

`serializeDocumentToDocx` walks the `Document` once before writing anything and
rejects it when any verbatim slot is not a single well-formed element. The
pieces already exist:

- `isSingleWellFormedElement(xml, expectedLocalName)`
  (`packages/core/src/docx/serializer/xmlUtils.ts`) is the predicate, already
  used for `rawPropertiesXml` in two serializers. It lives in `folio-core` and
  `docx-core` does not depend on it, so it moves to `docx-core` and the two
  existing callers import it from there.
- The expected local name is not always known — a preserved child is whatever
  the source had — so the check has two modes: named, for a slot whose element
  the model states (`sdtPr`, `tblGridChange`, `numberingChange`), and unnamed,
  for a capture, where the requirement is only "exactly one element, balanced,
  nothing after it".
- The XML resource limits in `packages/core/src/docx/xmlResourceLimits.ts`
  bound the walk: element count, attribute count and total bytes, charged
  against one budget for the whole document rather than per slot, so a
  thousand small slots cost what one large one does.
- The failure is a `TaggedError`, not a silent drop. A constructed document
  with unusable markup in it is a caller's bug, and returning a package that
  quietly lost a clause is the worse answer.

## The cost

- One traversal of the model per constructed save, over the same nodes the
  writer already visits. The parse-and-repack path does not pay it.
- `isSingleWellFormedElement` parses each slot's string. The legal-source
  compiler produces few verbatim slots, so the real cost is the resource-limit
  accounting rather than the parse.
- Moving the predicate into `docx-core` is a public-API move in both packages
  and needs an api-report regeneration and a changeset.
- The named/unnamed split has to be written down per slot, and it is a
  companion map over the slot list above: `as const satisfies Record<…>` so a
  new verbatim member cannot reach the writer without a decision.
