// Five fresh deletions; captured and loop-carried writes remain reachable.
export const flaggedSibling = (tr, schema) => {
  const mark = schema.marks.deletion.create({ revisionId: 1 });
  tr.addMark(1, 2, mark);
};
export const allowedSibling = (tr, schema) => {
  const mark = schema.marks.insertion.create({ revisionId: 2 });
  tr.addMark(1, 2, mark);
};
export const allowedParameter = (tr, mark) => tr.addMark(1, 2, mark);
export const nestedScopes = (tr, schema) => {
  const type = schema.marks.deletion;
  const mark = type.create({ revisionId: 3 });
  {
    const mark = schema.marks.insertion.create({ revisionId: 4 });
    tr.addMark(1, 2, mark);
  }
  const nested = () => tr.addMark(1, 2, mark);
  nested();
};
export const assignedInBlock = (tr, schema) => {
  let mark;
  {
    mark = schema.marks.deletion.create({ revisionId: 5 });
  }
  tr.addMark(1, 2, mark);
};
export const assignedAfterCapture = (tr, schema) => {
  let mark = schema.marks.insertion.create({ revisionId: 6 });
  const apply = () => tr.addMark(1, 2, mark);
  mark = schema.marks.deletion.create({ revisionId: 7 });
  apply();
};
export const loopCarriedDeletion = (tr, schema) => {
  let mark = schema.marks.insertion.create({ revisionId: 8 });
  for (let index = 0; index < 2; index++) {
    tr.addMark(1, 2, mark);
    mark = schema.marks.deletion.create({ revisionId: 9 });
  }
};
