import { forwardRef, type ReactNode } from "react";
import { IntlProvider, useTranslations } from "use-intl";

import {
  createDocumentWithText,
  createEmptyDocument,
  type CreateEmptyDocumentOptions,
} from "@stll/folio-core/utils/createDocument";

import { DocxEditor as FolioDocxEditor } from "../components/DocxEditor";
import type { DocxEditorProps, DocxEditorRef } from "../components/DocxEditor.props";
import { getFolioMessages } from "../i18n/messages";
export {
  renderAsync,
  type DocxEditorHandle,
  type EditorHandle,
  type RenderAsyncOptions,
} from "../renderAsync";

export const VERSION = "folio-compat-eigenpal";

type LegacyTranslations = {
  _lang?: string;
};

export type LocaleCode = "en" | "de" | "fr" | "he" | "hi" | "pl" | "pt-BR" | "tr" | "zh-CN";

export type LocaleProviderProps = {
  children?: ReactNode | undefined;
  i18n?: LegacyTranslations | undefined;
  locale?: string | undefined;
};

export type TFunction = (key: string, vars?: Record<string, string | number>) => string;

export type LegacyDocxEditorProps = DocxEditorProps & {
  agentPanel?: unknown;
  colorMode?: "light" | "dark" | "system";
  commentsSidebarOpen?: boolean;
  disableFindReplaceShortcuts?: boolean;
  documentName?: string;
  documentNameEditable?: boolean;
  externalContent?: boolean;
  externalPlugins?: unknown[];
  i18n?: LegacyTranslations;
  locale?: string;
  onCommentAdd?: unknown;
  onCommentDelete?: unknown;
  onCommentReply?: unknown;
  onCommentResolve?: unknown;
  onCommentsSidebarOpenChange?: unknown;
  onDocumentNameChange?: unknown;
  onOpen?: unknown;
  onRenderedDomContextReady?: unknown;
  pluginOverlays?: ReactNode;
  pluginRenderedDomContext?: unknown;
  pluginSidebarItems?: unknown[];
  printOptions?: unknown;
  renderLogo?: unknown;
  renderTitleBarRight?: unknown;
  showFileOpen?: boolean;
  showHelpMenu?: boolean;
  showOutlineButton?: boolean;
  watermarkPresets?: readonly string[];
};

export type { CreateEmptyDocumentOptions, DocxEditorRef };
export { createDocumentWithText, createEmptyDocument };

export const en = { _lang: "en" } satisfies LegacyTranslations;
export const de = { _lang: "de" } satisfies LegacyTranslations;
export const fr = { _lang: "fr" } satisfies LegacyTranslations;
export const he = { _lang: "he" } satisfies LegacyTranslations;
export const hi = { _lang: "hi" } satisfies LegacyTranslations;
export const pl = { _lang: "pl" } satisfies LegacyTranslations;
export const ptBR = { _lang: "pt-BR" } satisfies LegacyTranslations;
export const tr = { _lang: "tr" } satisfies LegacyTranslations;
export const zhCN = { _lang: "zh-CN" } satisfies LegacyTranslations;
export const locales: Record<LocaleCode, LegacyTranslations> = {
  en,
  de,
  fr,
  he,
  hi,
  pl,
  "pt-BR": ptBR,
  tr,
  "zh-CN": zhCN,
};

export const LocaleProvider = ({ children, i18n, locale }: LocaleProviderProps) => {
  const resolvedLocale = resolveLocale(locale, i18n);
  return (
    <IntlProvider locale={resolvedLocale} messages={getFolioMessages(resolvedLocale)}>
      {children}
    </IntlProvider>
  );
};

export const useTranslation = (): { t: TFunction } => {
  const t = useTranslations("folio");
  return { t: t as unknown as TFunction };
};

export const DocxEditor = forwardRef<DocxEditorRef, LegacyDocxEditorProps>(
  (
    {
      agentPanel: _agentPanel,
      colorMode: _colorMode,
      commentsSidebarOpen: _commentsSidebarOpen,
      disableFindReplaceShortcuts: _disableFindReplaceShortcuts,
      documentName: _documentName,
      documentNameEditable: _documentNameEditable,
      externalContent: _externalContent,
      externalPlugins: _externalPlugins,
      i18n,
      locale,
      onCommentAdd: _onCommentAdd,
      onCommentDelete: _onCommentDelete,
      onCommentReply: _onCommentReply,
      onCommentResolve: _onCommentResolve,
      onCommentsSidebarOpenChange: _onCommentsSidebarOpenChange,
      onDocumentNameChange: _onDocumentNameChange,
      onOpen: _onOpen,
      onRenderedDomContextReady: _onRenderedDomContextReady,
      pluginOverlays: _pluginOverlays,
      pluginRenderedDomContext: _pluginRenderedDomContext,
      pluginSidebarItems: _pluginSidebarItems,
      printOptions: _printOptions,
      renderLogo: _renderLogo,
      renderTitleBarRight: _renderTitleBarRight,
      showFileOpen: _showFileOpen,
      showHelpMenu: _showHelpMenu,
      showOutlineButton: _showOutlineButton,
      watermarkPresets: _watermarkPresets,
      ...props
    },
    ref,
  ) => (
    <LocaleProvider i18n={i18n} locale={locale}>
      <FolioDocxEditor {...props} ref={ref} />
    </LocaleProvider>
  ),
);

DocxEditor.displayName = "EigenpalCompatDocxEditor";

const resolveLocale = (locale: string | undefined, i18n: LegacyTranslations | undefined): string =>
  locale ?? i18n?._lang ?? "en";
