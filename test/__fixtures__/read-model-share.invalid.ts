// Fixture for `folio-read-model/model-share-owner`. Only the save shares a
// read-only value into the model it builds; both calls here are refused.

import type { TabStop } from "@stll/folio-core/types/document";
import * as Model from "@stll/folio-core/types/documentModel";
import {
  type DeepReadonly,
  shareWithModel,
} from "@stll/folio-core/types/documentModel";

export const direct = (tabs: DeepReadonly<TabStop[]>): TabStop[] =>
  shareWithModel<TabStop[]>(tabs);

export const qualified = (tabs: DeepReadonly<TabStop[]>): TabStop[] =>
  Model.shareWithModel<TabStop[]>(tabs);
