<script setup lang="ts">
// Fixture for `folio-read-model/no-model-document-cast` in a single-file
// component: the rule reads a component's script block as it reads a module.

import type { Document } from "@stll/folio-core/types/document";
import type { ReadonlyDocument } from "@stll/folio-core/types/documentModel";

const props = defineProps<{ model: ReadonlyDocument }>();

const direct = (): Document => props.model as Document;
const throughUnknown = (): Document => props.model as unknown as Document;

defineExpose({ direct, throughUnknown });
</script>

<template>
  <div />
</template>
