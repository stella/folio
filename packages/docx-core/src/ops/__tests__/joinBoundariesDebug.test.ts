import { test } from "bun:test";

import type { BlockContent, Document } from "../../model/document";
import { applyDocumentOp } from "../apply";
import { DOCUMENT_OP_TYPES, OP_STORIES, REVISION_DECISIONS } from "../types";

// Scratch debug: the tracked document of L3 seed -1270878361.
const tracked: BlockContent[] = [
  {
    "type": "paragraph",
    "content": [
      {
        "type": "insertion",
        "info": {
          "id": 303650,
          "author": "C",
          "date": "2026-03-04T05:00:02Z"
        },
        "content": [
          {
            "type": "run",
            "content": [
              {
                "type": "text",
                "text": "x"
              }
            ]
          }
        ],
        "resolutionJoins": {
          "before": 0,
          "after": 3,
          "remove": 0
        }
      },
      {
        "type": "insertion",
        "info": {
          "id": 303649,
          "author": "C",
          "date": "2026-03-04T05:00:01Z"
        },
        "content": [
          {
            "type": "run",
            "content": [
              {
                "type": "text",
                "text": "x"
              }
            ]
          }
        ],
        "resolutionJoins": {
          "before": 0,
          "after": 0,
          "remove": 0
        }
      }
    ],
    "paraId": "00000001"
  },
  {
    "type": "paragraph",
    "paraId": "10000000",
    "content": [
      {
        "type": "run",
        "content": [
          {
            "type": "text",
            "text": "a"
          }
        ]
      },
      {
        "type": "insertion",
        "info": {
          "id": 1,
          "author": "A"
        },
        "content": [
          {
            "type": "bookmarkEnd",
            "id": 3
          }
        ]
      }
    ],
    "pPrMark": {
      "kind": "ins",
      "info": {
        "id": 2,
        "author": "C",
        "date": "2026-03-04T05:00:00Z"
      },
      "resolutionJoin": 1
    }
  },
  {
    "type": "paragraph",
    "content": [
      {
        "type": "insertion",
        "info": {
          "id": 303648,
          "author": "A"
        },
        "content": [
          {
            "type": "run",
            "content": [
              {
                "type": "text",
                "text": "a"
              }
            ]
          }
        ]
      },
      {
        "type": "inlineSdt",
        "properties": {
          "sdtType": "richText",
          "id": 1,
          "tag": "clause"
        },
        "content": []
      }
    ],
    "sectionProperties": {
      "pageWidth": 11906
    },
    "paraId": "00000002"
  }
];
const source: BlockContent[] = [
  {
    "type": "paragraph",
    "content": [],
    "paraId": "00000001"
  },
  {
    "type": "paragraph",
    "content": [
      {
        "type": "run",
        "content": [
          {
            "type": "text",
            "text": "a"
          }
        ]
      },
      {
        "type": "insertion",
        "info": {
          "id": 1,
          "author": "A"
        },
        "content": [
          {
            "type": "bookmarkEnd",
            "id": 3
          },
          {
            "type": "run",
            "content": [
              {
                "type": "text",
                "text": "a"
              }
            ]
          }
        ]
      },
      {
        "type": "inlineSdt",
        "properties": {
          "sdtType": "richText",
          "id": 1,
          "tag": "clause"
        },
        "content": []
      }
    ],
    "sectionProperties": {
      "pageWidth": 11906
    },
    "paraId": "00000002"
  }
];

const reject = (document: Document, ids: number[]): Document => {
  const applied = applyDocumentOp(document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds: ids,
    decision: REVISION_DECISIONS.REJECT,
  });
  if (applied.isErr()) throw applied.error;
  return applied.value.document;
};

test("debug join boundaries seed -1270878361", () => {
  const document: Document = { package: { document: { content: structuredClone(tracked) } } };
  for (const ids of [[2], [303649], [303650], [2, 303649, 303650]]) {
    globalThis.__folioJoinDebug = true;
    const result = reject(structuredClone(document), ids);
    globalThis.__folioJoinDebug = false;
    console.error("DEBUGRESULT", JSON.stringify(ids), JSON.stringify(result.package.document.content));
  }
  console.error("DEBUGSOURCE", JSON.stringify(source));
});

declare global {
  var __folioJoinDebug: boolean | undefined;
}
