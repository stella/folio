const deletion = schema.marks["deletion"].create({ revisionId: 10 });
addTrackedDeletionMark({ tr, from: 1, to: 10, mark: deletion });
tr.addMark(1, 10, schema.marks["insertion"].create({ revisionId: 11 }));
// Restoring an existing revision's metadata is not a fresh deletion.
tr.addMark(1, 10, existing.type.create({ ...existing.attrs, author: "Reviewer" }));
