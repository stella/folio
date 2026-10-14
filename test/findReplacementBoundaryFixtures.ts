import type {
  Document,
  ParagraphContent,
  Run,
  RunContent,
  TrackedRunContent,
} from "@stll/folio-core/types/document";

const textRun = (text: string): Run => ({ type: "run", content: [{ type: "text", text }] });

type SourceDisposition = {
  status: "search-text" | "container";
};
type RunSourceCensus = {
  [Kind in RunContent["type"]]:
    | SourceDisposition
    | { status: "field-structure"; source: RunContent & { type: Kind } }
    | {
        status: "zero-text";
        content: { type: "run"; content: (RunContent & { type: Kind })[] }[];
      };
};
type ParagraphSourceCensus = {
  [Kind in ParagraphContent["type"]]:
    | SourceDisposition
    | {
        status: "structural-marker";
        source: ParagraphContent & { type: Kind };
        companions: ParagraphContent[];
      }
    | {
        status: "zero-text";
        content: (TrackedRunContent & { type: Kind })[];
      };
};

export const RUN_SEARCH_SOURCE_CENSUS = {
  text: { status: "search-text" },
  tab: { status: "search-text" },
  break: { status: "search-text" },
  symbol: { status: "search-text" },
  softHyphen: { status: "search-text" },
  noBreakHyphen: { status: "search-text" },
  // Standalone markers are consumed by the field parser; the complete cached
  // field fixtures below exercise their surviving PM carrier.
  fieldChar: {
    status: "field-structure",
    source: { type: "fieldChar", charType: "begin", fldLock: true, dirty: false },
  },
  instrText: { status: "field-structure", source: { type: "instrText", text: "REF target" } },
  footnoteRef: {
    status: "zero-text",
    content: [{ type: "run", content: [{ type: "footnoteRef", id: 37, customMarkFollows: true }] }],
  },
  endnoteRef: {
    status: "zero-text",
    content: [{ type: "run", content: [{ type: "endnoteRef", id: 41, customMarkFollows: true }] }],
  },
  noteMarker: {
    status: "zero-text",
    content: [{ type: "run", content: [{ type: "noteMarker", kind: "footnote" }] }],
  },
  renderedPageBreak: {
    status: "zero-text",
    content: [{ type: "run", content: [{ type: "renderedPageBreak" }] }],
  },
  preservedXml: {
    status: "zero-text",
    content: [
      {
        type: "run",
        content: [
          {
            type: "preservedXml",
            text: "",
            xml: '<w:separator xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
          },
        ],
      },
    ],
  },
  drawing: {
    status: "zero-text",
    content: [
      {
        type: "run",
        content: [
          {
            type: "drawing",
            image: {
              type: "image",
              size: { width: 914400, height: 914400 },
              wrap: { type: "inline" },
              alt: "Preserved illustration",
            },
          },
        ],
      },
    ],
  },
  shape: {
    status: "zero-text",
    content: [
      {
        type: "run",
        content: [
          {
            type: "shape",
            shape: {
              type: "shape",
              shapeType: "rect",
              id: "Shape 37",
              size: { width: 914400, height: 914400 },
            },
          },
        ],
      },
    ],
  },
} satisfies RunSourceCensus;

export const PARAGRAPH_SEARCH_SOURCE_CENSUS = {
  run: { status: "container" },
  hyperlink: { status: "container" },
  inlineSdt: { status: "container" },
  inlineWrapper: { status: "container" },
  insertion: { status: "container" },
  deletion: { status: "container" },
  moveFrom: { status: "container" },
  moveTo: { status: "container" },
  bookmarkStart: {
    status: "structural-marker",
    source: { type: "bookmarkStart", id: 71, name: "SearchBoundary" },
    companions: [{ type: "bookmarkEnd", id: 71 }],
  },
  bookmarkEnd: {
    status: "structural-marker",
    source: { type: "bookmarkEnd", id: 71 },
    companions: [{ type: "bookmarkStart", id: 71, name: "SearchBoundary" }],
  },
  commentRangeStart: {
    status: "structural-marker",
    source: { type: "commentRangeStart", id: 91 },
    companions: [{ type: "commentRangeEnd", id: 91 }],
  },
  commentRangeEnd: {
    status: "structural-marker",
    source: { type: "commentRangeEnd", id: 91 },
    companions: [{ type: "commentRangeStart", id: 91 }],
  },
  commentReference: {
    status: "structural-marker",
    source: { type: "commentReference", id: 91 },
    companions: [],
  },
  moveFromRangeStart: {
    status: "structural-marker",
    source: {
      type: "moveFromRangeStart",
      id: 83,
      name: "MovedSource",
      author: "Reviewer",
      date: "2026-10-14T10:00:00Z",
    },
    companions: [{ type: "moveFromRangeEnd", id: 83 }],
  },
  moveFromRangeEnd: {
    status: "structural-marker",
    source: { type: "moveFromRangeEnd", id: 83 },
    companions: [
      {
        type: "moveFromRangeStart",
        id: 83,
        name: "MovedSource",
        author: "Reviewer",
        date: "2026-10-14T10:00:00Z",
      },
    ],
  },
  moveToRangeStart: {
    status: "structural-marker",
    source: {
      type: "moveToRangeStart",
      id: 85,
      name: "MovedTarget",
      author: "Reviewer",
      date: "2026-10-14T10:00:00Z",
    },
    companions: [{ type: "moveToRangeEnd", id: 85 }],
  },
  moveToRangeEnd: {
    status: "structural-marker",
    source: { type: "moveToRangeEnd", id: 85 },
    companions: [
      {
        type: "moveToRangeStart",
        id: 85,
        name: "MovedTarget",
        author: "Reviewer",
        date: "2026-10-14T10:00:00Z",
      },
    ],
  },
  simpleField: {
    status: "zero-text",
    content: [{ type: "simpleField", fieldType: "REF", instruction: "REF target", content: [] }],
  },
  complexField: {
    status: "zero-text",
    content: [
      {
        type: "complexField",
        fieldType: "REF",
        instruction: "REF target",
        fieldCode: [],
        fieldResult: [],
      },
    ],
  },
  mathEquation: {
    status: "zero-text",
    content: [
      {
        type: "mathEquation",
        display: "inline",
        ommlXml: '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"/>',
      },
    ],
  },
  preservedInline: {
    status: "zero-text",
    content: [
      {
        type: "preservedInline",
        text: "",
        xml: '<w:proofErr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:type="spellStart"/>',
      },
    ],
  },
} satisfies ParagraphSourceCensus;

const zeroTextBoundaryFixtures = () => {
  const fixtures: Record<string, TrackedRunContent[]> = {};
  for (const census of [RUN_SEARCH_SOURCE_CENSUS, PARAGRAPH_SEARCH_SOURCE_CENSUS]) {
    for (const [kind, disposition] of Object.entries(census)) {
      if (disposition.status === "zero-text") fixtures[kind] = disposition.content;
    }
  }
  return fixtures;
};

export const ZERO_SEARCH_TEXT_BOUNDARY_FIXTURES = zeroTextBoundaryFixtures();

const sourceZeroTextBoundaryFixtures = () => {
  const fixtures: Record<string, TrackedRunContent[]> = { ...ZERO_SEARCH_TEXT_BOUNDARY_FIXTURES };
  for (const [kind, disposition] of Object.entries(RUN_SEARCH_SOURCE_CENSUS)) {
    if (disposition.status === "field-structure")
      fixtures[kind] = [{ type: "run", content: [disposition.source] }];
  }
  for (const [kind, disposition] of Object.entries(PARAGRAPH_SEARCH_SOURCE_CENSUS)) {
    if (disposition.status !== "structural-marker") continue;
    const content = disposition.source.type.endsWith("End")
      ? [...disposition.companions, disposition.source]
      : [disposition.source, ...disposition.companions];
    // Comment markers are legal inside this transparent wrapper, including
    // when the wrapper itself sits in a tracked deletion or move.
    fixtures[kind] = [
      { type: "inlineWrapper", kind: "customXml", element: "SearchBoundary", content },
    ];
  }
  return fixtures;
};

export const SOURCE_ZERO_TEXT_BOUNDARY_FIXTURES = sourceZeroTextBoundaryFixtures();

export const CONDITIONAL_SOURCE_ZERO_TEXT_BOUNDARY_FIXTURES = {
  undecodableSymbol: [
    { type: "run", content: [{ type: "symbol", font: "Wingdings", char: "invalid" }] },
  ],
  pageBreak: [{ type: "run", content: [{ type: "break", breakType: "page" }] }],
} satisfies Record<string, TrackedRunContent[]>;

export const buildBoundaryDocument = (content: ParagraphContent[]) =>
  ({
    package: {
      document: {
        content: [{ type: "paragraph", paraId: "4AB712CD", content }],
        comments: [
          {
            id: 91,
            author: "Reviewer",
            content: [{ type: "paragraph", content: [textRun("Comment record")] }],
          },
        ],
      },
      footnotes: [
        {
          type: "footnote",
          id: 37,
          content: [{ type: "paragraph", content: [textRun("Footnote record")] }],
        },
      ],
      endnotes: [
        {
          type: "endnote",
          id: 41,
          content: [{ type: "paragraph", content: [textRun("Endnote record")] }],
        },
      ],
    },
  }) satisfies Document;

export const surroundHiddenBoundary = (
  content: TrackedRunContent[],
  revisionType: "deletion" | "moveFrom" = "deletion",
): ParagraphContent[] => [
  textRun("A"),
  {
    type: revisionType,
    info: { id: 1701037, author: "Reviewer", initials: "R", date: "2026-10-14T10:00:00Z" },
    content,
  },
  textRun("B"),
];
