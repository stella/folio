export const ownedDeletion = (tr, schema) => {
  const mark = schema.marks.deletion.create({ revisionId: 1 });
  addTrackedDeletionMark({ tr, from: 1, to: 2, mark });
};
export const siblingInsertion = (tr, schema) => {
  const mark = schema.marks.insertion.create({ revisionId: 2 });
  tr.addMark(1, 2, mark);
};
export const parameterShadow = (tr, mark) => tr.addMark(1, 2, mark);
export const typeShadow = (tr, schema) => {
  const type = schema.marks.deletion;
  const mark = type.create({ revisionId: 3 });
  addTrackedDeletionMark({ tr, from: 1, to: 2, mark });
  {
    const type = schema.marks.insertion;
    tr.addMark(1, 2, type.create({ revisionId: 4 }));
  }
};
export const laterDeletion = (tr, schema) => {
  let mark = schema.marks.insertion.create({ revisionId: 5 });
  tr.addMark(1, 2, mark);
  mark = schema.marks.deletion.create({ revisionId: 6 });
  addTrackedDeletionMark({ tr, from: 1, to: 2, mark });
};
