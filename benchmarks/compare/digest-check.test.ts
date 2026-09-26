import { expect, test } from "bun:test";

import { checkRecordedDigests } from "./digest-check";

const alpha = { buffer: "base", changes: "changes" };
const beta = { buffer: "next", changes: "next-changes" };

test("missing baselines cannot pass the digest gate", () => {
  expect(checkRecordedDigests([{ id: "plain/s/light", digests: alpha }], {})).toEqual({
    drifted: [],
    missing: ["plain/s/light"],
    checked: 0,
  });
});

test("digest gate reports missing and drifted configurations independently", () => {
  expect(
    checkRecordedDigests(
      [
        { id: "plain/s/light", digests: alpha },
        { id: "plain/s/heavy", digests: beta },
        { id: "plain/s/churn", digests: beta },
      ],
      {
        "plain/s/light": alpha,
        "plain/s/heavy": alpha,
      },
    ),
  ).toEqual({ drifted: ["plain/s/heavy"], missing: ["plain/s/churn"], checked: 2 });
});
