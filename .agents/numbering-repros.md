## Appendix: repro lines (seed, fixture / mode)

Lines marked SEED-ONLY had a truncated replay in the report; try the generic form shown and, if it does not fail, mark the issue Related. Older lines without a trailing file argument run every scenario file filtered by name; append `-- fuzz.test.ts` to shorten them. Lines replaying `FOLIO_SCENARIO_FLOW` are already minimised flows: they are the corpus files to pin.

- #1917 trackedChanges / direct (seeds 1502256799):
  `FOLIO_SCENARIO_FLOW='{"version":1,"kind":"random","generation":"targeted","fixture":"trackedChanges","mode":"direct","seed":1502256799,"steps":[{"action":"mistake","seed":2502587051}],"swarm":["commentOnRange","insertAfterBlock","replaceBlock","deleteBlock","setBlockParagraphProperties","insertTable","insertTableRow","insertTableColumn","deleteTable"],"origin":"shrunk from random flow seed 1502256799"}' bun scripts/consumer-scenarios.ts --only '^replay FOLIO_SCENARIO_FLOW' -- flow-corpus.test.ts`
- #1914 directNumbered / direct (seeds 1502256264, 1502256745):
  `FOLIO_SCENARIO_FLOW='{"version":1,"kind":"collisions","generation":"targeted","fixture":"directNumbered","mode":"direct","seed":1502256264,"steps":[{"action":"mistake","seed":4270211132}],"swarm":["replaceInBlock","deleteBlock","splitBlock","mergeBlockWithNext","commentOnBlock","insertTable","insertTableColumn"],"origin":"shrunk from collisions flow seed 1502256264"}' bun scripts/consumer-scenarios.ts --only '^replay FOLIO_SCENARIO_FLOW' -- flow-corpus.test.ts`
- #1911 stories / direct (seeds 1502256267, 1346488458):
  `FOLIO_SCENARIO_SEED=1346488332 FOLIO_SCENARIO_FUZZ_STEPS=16 FOLIO_SCENARIO_FUZZ_RUNS=127 FOLIO_SCENARIO_COLLISION_RUNS=0  bun scripts/consumer-scenarios.ts --only '^fuzz run 126 \('`
- #1876 trackedChanges / direct (seeds 740097592) SEED-ONLY:
  `FOLIO_SCENARIO_SEED=740097592 FOLIO_SCENARIO_FUZZ_STEPS=20 FOLIO_SCENARIO_FUZZ_RUNS=1 FOLIO_SCENARIO_COLLISION_RUNS=0 FOLIO_SCENARIO_SWARM=1 bun scripts/consumer-scenarios.ts --only '^fuzz run 0 \(' -- fuzz.test.ts`
- #1844 directNumbered / direct (seeds 1101252770):
  `FOLIO_SCENARIO_SEED=1101252625 FOLIO_SCENARIO_FUZZ_STEPS=16 FOLIO_SCENARIO_FUZZ_RUNS=0 FOLIO_SCENARIO_COLLISION_RUNS=146  bun scripts/consumer-scenarios.ts --only '^collision run 145 \('`
- #1837 lists / direct (seeds 1101252663):
  `FOLIO_SCENARIO_SEED=1101252625 FOLIO_SCENARIO_FUZZ_STEPS=16 FOLIO_SCENARIO_FUZZ_RUNS=0 FOLIO_SCENARIO_COLLISION_RUNS=39  bun scripts/consumer-scenarios.ts --only '^collision run 38 \('`
- #1830 trackedChanges / direct (seeds 1101252728):
  `FOLIO_SCENARIO_SEED=1101252625 FOLIO_SCENARIO_FUZZ_STEPS=16 FOLIO_SCENARIO_FUZZ_RUNS=104 FOLIO_SCENARIO_COLLISION_RUNS=0  bun scripts/consumer-scenarios.ts --only '^fuzz run 103 \('`
- #1771 stories / direct (seeds 746948075):
  `FOLIO_SCENARIO_SEED=746947873 FOLIO_SCENARIO_FUZZ_STEPS=16 FOLIO_SCENARIO_FUZZ_RUNS=203 FOLIO_SCENARIO_COLLISION_RUNS=0  bun scripts/consumer-scenarios.ts --only '^fuzz run 202 \('`
- #1766 styleNumbered / direct (seeds 746948022):
  `FOLIO_SCENARIO_SEED=746947873 FOLIO_SCENARIO_FUZZ_STEPS=16 FOLIO_SCENARIO_FUZZ_RUNS=150 FOLIO_SCENARIO_COLLISION_RUNS=0  bun scripts/consumer-scenarios.ts --only '^fuzz run 149 \('`
- #1745 directNumbered / direct (seeds 1809931248):
  `FOLIO_SCENARIO_SEED=1809931248 FOLIO_SCENARIO_FUZZ_STEPS=20 FOLIO_SCENARIO_FUZZ_RUNS=1 FOLIO_SCENARIO_COLLISION_RUNS=0 FOLIO_SCENARIO_SWARM=1 bun scripts/consumer-scenarios.ts --only '^fuzz run 0 \(' -- fuzz.test.ts`
- #1739 lists / direct (seeds 1809931261) SEED-ONLY:
  `FOLIO_SCENARIO_SEED=1809931261 FOLIO_SCENARIO_FUZZ_STEPS=20 FOLIO_SCENARIO_FUZZ_RUNS=1 FOLIO_SCENARIO_COLLISION_RUNS=0 FOLIO_SCENARIO_SWARM=1 bun scripts/consumer-scenarios.ts --only '^fuzz run 0 \(' -- fuzz.test.ts`
- #1735 stories / direct (seeds 1809930896, 1809931265):
  `FOLIO_SCENARIO_SEED=1809931265 FOLIO_SCENARIO_FUZZ_STEPS=20 FOLIO_SCENARIO_FUZZ_RUNS=0 FOLIO_SCENARIO_COLLISION_RUNS=1 FOLIO_SCENARIO_SWARM=1 bun scripts/consumer-scenarios.ts --only '^collision run 0 \(' -- fuzz.test.ts`
- #1913 directNumbered / tracked-changes (seeds 1502256833):
  `FOLIO_SCENARIO_FLOW='{"version":1,"kind":"random","generation":"targeted","fixture":"directNumbered","mode":"tracked-changes","seed":1502256833,"steps":[{"action":"mistake","seed":926786060}],"swarm":["replaceInBlock","formatRange","commentOnRange","insertBeforeBlock","replaceBlock","splitBlock","commentOnBlock","insertTable","insertTableRow","deleteTableRow","insertTableColumn"],"origin":"shrunk from random flow seed 1502256833"}' bun scripts/consumer-scenarios.ts --only '^replay FOLIO_SCENARIO_FLOW' -- flow-corpus.test.ts`
- #1895 emoji / tracked-changes (seeds 14491387):
  `FOLIO_SCENARIO_SEED=14491268 FOLIO_SCENARIO_FUZZ_STEPS=16 FOLIO_SCENARIO_FUZZ_RUNS=0 FOLIO_SCENARIO_COLLISION_RUNS=120  bun scripts/consumer-scenarios.ts --only '^collision run 119 \('`
- #1799 styleNumbered / tracked-changes (seeds 1339716102, 1339716619):
  `FOLIO_SCENARIO_FLOW='{"version":1,"kind":"collisions","generation":"targeted","fixture":"styleNumbered","mode":"tracked-changes","seed":1492413172,"steps":[{"action":"core batch","seed":476930320},{"action":"mistake","seed":1924848951},{"action":"story batch","seed":4137292766}],"swarm":["replaceInBlock","insertAfterBlock","insertBeforeBlock","replaceBlock","mergeBlockWithNext","setBlockParagraphProperties","deleteTableRow"],"origin":"unpin+insert of 7ab6cd0460ca596b"}' bun scripts/consumer-scenarios.ts --only '^replay FOLIO_SCENARIO_FLOW' -- flow-corpus.test.ts`
- #1772 directNumbered / tracked-changes (seeds 746948112):
  `FOLIO_SCENARIO_SEED=746947873 FOLIO_SCENARIO_FUZZ_STEPS=16 FOLIO_SCENARIO_FUZZ_RUNS=240 FOLIO_SCENARIO_COLLISION_RUNS=0  bun scripts/consumer-scenarios.ts --only '^fuzz run 239 \('`
- #1753 lists / tracked-changes (seeds 1809931245):
  `FOLIO_SCENARIO_SEED=1809931245 FOLIO_SCENARIO_FUZZ_STEPS=20 FOLIO_SCENARIO_FUZZ_RUNS=0 FOLIO_SCENARIO_COLLISION_RUNS=1 FOLIO_SCENARIO_SWARM=1 bun scripts/consumer-scenarios.ts --only '^collision run 0 \(' -- fuzz.test.ts`
- #1732 comments / tracked-changes (seeds 1809930875) SEED-ONLY:
  `FOLIO_SCENARIO_SEED=1809930875 FOLIO_SCENARIO_FUZZ_STEPS=20 FOLIO_SCENARIO_FUZZ_RUNS=1 FOLIO_SCENARIO_COLLISION_RUNS=0 FOLIO_SCENARIO_SWARM=1 bun scripts/consumer-scenarios.ts --only '^fuzz run 0 \(' -- fuzz.test.ts`
- #1731 stories / tracked-changes (seeds 1809930952):
  `FOLIO_SCENARIO_SEED=1809930952 FOLIO_SCENARIO_FUZZ_STEPS=20 FOLIO_SCENARIO_FUZZ_RUNS=1 FOLIO_SCENARIO_COLLISION_RUNS=0 FOLIO_SCENARIO_SWARM=1 bun scripts/consumer-scenarios.ts --only '^fuzz run 0 \(' -- fuzz.test.ts`
- #1730 directNumbered / tracked-changes (seeds 1809931253):
  `FOLIO_SCENARIO_SEED=1809931253 FOLIO_SCENARIO_FUZZ_STEPS=20 FOLIO_SCENARIO_FUZZ_RUNS=0 FOLIO_SCENARIO_COLLISION_RUNS=1 FOLIO_SCENARIO_SWARM=1 bun scripts/consumer-scenarios.ts --only '^collision run 0 \(' -- fuzz.test.ts`
- #1729 trackedChanges / tracked-changes (seeds 1809931153):
  `FOLIO_SCENARIO_SEED=1809931153 FOLIO_SCENARIO_FUZZ_STEPS=20 FOLIO_SCENARIO_FUZZ_RUNS=0 FOLIO_SCENARIO_COLLISION_RUNS=1 FOLIO_SCENARIO_SWARM=1 bun scripts/consumer-scenarios.ts --only '^collision run 0 \(' -- fuzz.test.ts`
- #1931 trackedChanges / suggested (seeds 1502256768, 1346488368):
  `FOLIO_SCENARIO_SEED=1346488332 FOLIO_SCENARIO_FUZZ_STEPS=16 FOLIO_SCENARIO_FUZZ_RUNS=0 FOLIO_SCENARIO_COLLISION_RUNS=37  bun scripts/consumer-scenarios.ts --only '^collision run 36 \('`
- #1886 directNumbered / suggested (seeds 14491290):
  `FOLIO_SCENARIO_SEED=14491268 FOLIO_SCENARIO_FUZZ_STEPS=16 FOLIO_SCENARIO_FUZZ_RUNS=23 FOLIO_SCENARIO_COLLISION_RUNS=0  bun scripts/consumer-scenarios.ts --only '^fuzz run 22 \('`
- #1868 trackedChanges / suggested (seeds 740097347, 740097692):
  `FOLIO_SCENARIO_SEED=740097692 FOLIO_SCENARIO_FUZZ_STEPS=20 FOLIO_SCENARIO_FUZZ_RUNS=0 FOLIO_SCENARIO_COLLISION_RUNS=1 FOLIO_SCENARIO_SWARM=1 bun scripts/consumer-scenarios.ts --only '^collision run 0 \(' -- fuzz.test.ts`
- #1819 directNumbered / suggested (seeds 1339716047, 1339716277):
  `FOLIO_SCENARIO_SEED=1339716277 FOLIO_SCENARIO_FUZZ_STEPS=20 FOLIO_SCENARIO_FUZZ_RUNS=1 FOLIO_SCENARIO_COLLISION_RUNS=0 FOLIO_SCENARIO_SWARM=1 bun scripts/consumer-scenarios.ts --only '^fuzz run 0 \(' -- fuzz.test.ts`
- #1804 stories / suggested (seeds 1339716106):
  `FOLIO_SCENARIO_SEED=1339716106 FOLIO_SCENARIO_FUZZ_STEPS=20 FOLIO_SCENARIO_FUZZ_RUNS=0 FOLIO_SCENARIO_COLLISION_RUNS=1 FOLIO_SCENARIO_SWARM=1 bun scripts/consumer-scenarios.ts --only '^collision run 0 \(' -- fuzz.test.ts`
- #1774 stories / suggested (seeds 746947993):
  `FOLIO_SCENARIO_SEED=746947873 FOLIO_SCENARIO_FUZZ_STEPS=16 FOLIO_SCENARIO_FUZZ_RUNS=0 FOLIO_SCENARIO_COLLISION_RUNS=121  bun scripts/consumer-scenarios.ts --only '^collision run 120 \('`
- #1761 styleNumbered / suggested (seeds 746947935):
  `FOLIO_SCENARIO_SEED=746947873 FOLIO_SCENARIO_FUZZ_STEPS=16 FOLIO_SCENARIO_FUZZ_RUNS=63 FOLIO_SCENARIO_COLLISION_RUNS=0  bun scripts/consumer-scenarios.ts --only '^fuzz run 62 \('`
- #1738 lists / suggested (seeds 1809930854, 1809930955):
  `FOLIO_SCENARIO_SEED=1809930955 FOLIO_SCENARIO_FUZZ_STEPS=20 FOLIO_SCENARIO_FUZZ_RUNS=0 FOLIO_SCENARIO_COLLISION_RUNS=1 FOLIO_SCENARIO_SWARM=1 bun scripts/consumer-scenarios.ts --only '^collision run 0 \(' -- fuzz.test.ts`
- #1877 stories / tracked-changes (seeds 740097530):
  `FOLIO_SCENARIO_SEED=740097530 FOLIO_SCENARIO_FUZZ_STEPS=20 FOLIO_SCENARIO_FUZZ_RUNS=0 FOLIO_SCENARIO_COLLISION_RUNS=1 FOLIO_SCENARIO_SWARM=1 bun scripts/consumer-scenarios.ts --only '^collision run 0 \(' -- fuzz.test.ts`
- #1871 lists / tracked-changes (seeds 740097286):
  `FOLIO_SCENARIO_FLOW='{"version":1,"kind":"collisions","generation":"targeted","fixture":"lists","mode":"tracked-changes","seed":1026399033,"steps":[{"action":"core batch","seed":2678616336},{"action":"core batch","seed":668579786},{"action":"mistake","seed":4267770215},{"action":"new reviewer","seed":1172533062},{"action":"reply and resolve","seed":3004098875},{"action":"story batch","seed":540697392},{"action":"suggest_changes","seed":2308322008},{"action":"core batch","seed":348873801},{"action":"reject one","seed":93109013},{"action":"save and reopen","seed":3756240639}],"weights":{"version":1,"cells":{},"operations":[]},"origin":"delete+reseed+splice of 686a2692acb81e00"}' bun scripts/consumer-scenarios.ts --only '^replay FOLIO_SCENARIO_FLOW' -- flow-corpus.test.ts`
- #1800 trackedChanges / tracked-changes (seeds 1339715958):
  `FOLIO_SCENARIO_SEED=1339715958 FOLIO_SCENARIO_FUZZ_STEPS=20 FOLIO_SCENARIO_FUZZ_RUNS=0 FOLIO_SCENARIO_COLLISION_RUNS=1 FOLIO_SCENARIO_SWARM=1 bun scripts/consumer-scenarios.ts --only '^collision run 0 \(' -- fuzz.test.ts`
- #1770 lists / tracked-changes (seeds 746948072):
  `FOLIO_SCENARIO_SEED=746947873 FOLIO_SCENARIO_FUZZ_STEPS=16 FOLIO_SCENARIO_FUZZ_RUNS=200 FOLIO_SCENARIO_COLLISION_RUNS=0  bun scripts/consumer-scenarios.ts --only '^fuzz run 199 \('`
- #1768 trackedChanges / tracked-changes (seeds 746948060):
  `FOLIO_SCENARIO_SEED=746947873 FOLIO_SCENARIO_FUZZ_STEPS=16 FOLIO_SCENARIO_FUZZ_RUNS=188 FOLIO_SCENARIO_COLLISION_RUNS=0  bun scripts/consumer-scenarios.ts --only '^fuzz run 187 \('`
- #1839 directNumbered / tracked-changes (seeds 1101252671):
  `FOLIO_SCENARIO_SEED=1101252625 FOLIO_SCENARIO_FUZZ_STEPS=16 FOLIO_SCENARIO_FUZZ_RUNS=0 FOLIO_SCENARIO_COLLISION_RUNS=47  bun scripts/consumer-scenarios.ts --only '^collision run 46 \('`
- #1815 styleNumbered / tracked-changes (seeds 1339716569):
  `FOLIO_SCENARIO_SEED=1339716569 FOLIO_SCENARIO_FUZZ_STEPS=20 FOLIO_SCENARIO_FUZZ_RUNS=1 FOLIO_SCENARIO_COLLISION_RUNS=0 FOLIO_SCENARIO_SWARM=1 bun scripts/consumer-scenarios.ts --only '^fuzz run 0 \(' -- fuzz.test.ts`
- #1807 lists / tracked-changes (seeds 1339715829, 1339715970, 1339716074):
  `FOLIO_SCENARIO_FLOW='{"version":1,"kind":"collisions","generation":"targeted","fixture":"lists","mode":"tracked-changes","seed":1026399033,"steps":[{"action":"suggest_changes","seed":962457836},{"action":"core batch","seed":1778540797},{"action":"new reviewer","seed":1664497593}],"weights":{"version":1,"cells":{},"operations":[]},"origin":"unpin+unpin of 7f3a3eb320d80fa7"}' bun scripts/consumer-scenarios.ts --only '^replay FOLIO_SCENARIO_FLOW' -- flow-corpus.test.ts`
- #1760 trackedChanges / tracked-changes (seeds 746947921):
  `FOLIO_SCENARIO_SEED=746947873 FOLIO_SCENARIO_FUZZ_STEPS=16 FOLIO_SCENARIO_FUZZ_RUNS=49 FOLIO_SCENARIO_COLLISION_RUNS=0  bun scripts/consumer-scenarios.ts --only '^fuzz run 48 \('`
- #1726 trackedChanges / tracked-changes (seeds 1809930759):
  `FOLIO_SCENARIO_SEED=1809930759 FOLIO_SCENARIO_FUZZ_STEPS=20 FOLIO_SCENARIO_FUZZ_RUNS=0 FOLIO_SCENARIO_COLLISION_RUNS=1 FOLIO_SCENARIO_SWARM=1 bun scripts/consumer-scenarios.ts --only '^collision run 0 \(' -- fuzz.test.ts`
- #1863 directNumbered / tracked-changes (seeds 740097491):
  `FOLIO_SCENARIO_SEED=740097491 FOLIO_SCENARIO_FUZZ_STEPS=20 FOLIO_SCENARIO_FUZZ_RUNS=1 FOLIO_SCENARIO_COLLISION_RUNS=0 FOLIO_SCENARIO_SWARM=1 bun scripts/consumer-scenarios.ts --only '^fuzz run 0 \(' -- fuzz.test.ts`
- #1767 directNumbered / tracked-changes (seeds 746948031):
  `FOLIO_SCENARIO_SEED=746947873 FOLIO_SCENARIO_FUZZ_STEPS=16 FOLIO_SCENARIO_FUZZ_RUNS=159 FOLIO_SCENARIO_COLLISION_RUNS=0  bun scripts/consumer-scenarios.ts --only '^fuzz run 158 \('`
- #1746 directNumbered / tracked-changes (seeds 1809930746, 1809931056, 1809931334):
  `FOLIO_SCENARIO_SEED=1809931334 FOLIO_SCENARIO_FUZZ_STEPS=20 FOLIO_SCENARIO_FUZZ_RUNS=1 FOLIO_SCENARIO_COLLISION_RUNS=0 FOLIO_SCENARIO_SWARM=1 bun scripts/consumer-scenarios.ts --only '^fuzz run 0 \(' -- fuzz.test.ts`
- #1838 directNumbered / direct (seeds 1101252664):
  `FOLIO_SCENARIO_SEED=1101252625 FOLIO_SCENARIO_FUZZ_STEPS=16 FOLIO_SCENARIO_FUZZ_RUNS=0 FOLIO_SCENARIO_COLLISION_RUNS=40  bun scripts/consumer-scenarios.ts --only '^collision run 39 \('`
- #1817 comments / direct (seeds 1339716540):
  `FOLIO_SCENARIO_SEED=1339716540 FOLIO_SCENARIO_FUZZ_STEPS=20 FOLIO_SCENARIO_FUZZ_RUNS=0 FOLIO_SCENARIO_COLLISION_RUNS=1 FOLIO_SCENARIO_SWARM=1 bun scripts/consumer-scenarios.ts --only '^collision run 0 \(' -- fuzz.test.ts`
