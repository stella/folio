// Exactly three fresh deletions; sibling/parameter/block names cannot contaminate calls.
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
