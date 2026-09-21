---
"@stll/folio-core": patch
---

Check that a resolved story persisted without checking an id the save may not write.

`resolveReviewedStory` records what the resolved story should look like in the package, and the check after the save compares the reopened story against it — block by block and through the reading text, which carries each block's id. folio deterministically derives an id from the text and ordinal of every paragraph that arrives without a `w14:paraId`, deriving it again whenever an id-less paragraph is parsed. The selective patch deliberately writes none of them back, which is what keeps a package from a producer that writes no ids id-less and an edit to it local, so the reopened story named its paragraphs differently and `toBuffer` refused a story that had persisted exactly: `FolioResolvedStorySerializationError` with `text-projection` and `block-projection`, for any document whose paragraphs carry no ids.

The comparison now stands a minted id down to its position, on both sides and at the same positions, so it no longer depends on the id or on which save path ran — a full repack writes the model's ids while the selective patch does not. An authored id is compared exactly as before, and a block that moved still fails.
