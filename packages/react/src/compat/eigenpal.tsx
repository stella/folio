import {
  forwardRef,
  type ForwardRefExoticComponent,
  type JSX,
  type ReactNode,
  type RefAttributes,
  useEffect,
  useState,
} from "react";

import { IntlProvider, useTranslations } from "use-intl";

import {
  getLoadedFolioMessages,
  loadFolioMessages,
} from "@stll/folio-core/i18n/loadMessages";
import type { FolioMessages } from "@stll/folio-core/i18n/messages";
import { messages as english } from "@stll/folio-core/i18n/messages/en";
import { DEFAULT_KEYBOARD_SHORTCUT_SCOPE } from "@stll/folio-core/managers/editorShortcuts";
import {
  createDocumentWithText,
  createEmptyDocument,
  type CreateEmptyDocumentOptions,
} from "@stll/folio-core/utils/createDocument";

import { DocxEditor as FolioDocxEditor } from "../components/DocxEditor";
import type {
  DocxEditorProps,
  DocxEditorRef,
} from "../components/DocxEditor.props";
import { EditorLoadingIndicator } from "../components/DocxEditorHelpers";

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

export type LocaleCode =
  "en" | "de" | "fr" | "he" | "hi" | "pl" | "pt-BR" | "tr" | "zh-CN";

export type LocaleProviderProps = {
  children?: ReactNode | undefined;
  i18n?: LegacyTranslations | undefined;
  locale?: string | undefined;
  onError?: ((error: Error) => void) | undefined;
};

export type TFunction = (
  key: string,
  vars?: Record<string, string | number>,
) => string;

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

export const en: {
  _lang: string;
} = { _lang: "en" } satisfies LegacyTranslations;
export const de: {
  _lang: string;
} = { _lang: "de" } satisfies LegacyTranslations;
export const fr: {
  _lang: string;
} = { _lang: "fr" } satisfies LegacyTranslations;
export const he: {
  _lang: string;
} = { _lang: "he" } satisfies LegacyTranslations;
export const hi: {
  _lang: string;
} = { _lang: "hi" } satisfies LegacyTranslations;
export const pl: {
  _lang: string;
} = { _lang: "pl" } satisfies LegacyTranslations;
export const ptBR: {
  _lang: string;
} = { _lang: "pt-BR" } satisfies LegacyTranslations;
export const tr: {
  _lang: string;
} = { _lang: "tr" } satisfies LegacyTranslations;
export const zhCN: {
  _lang: string;
} = { _lang: "zh-CN" } satisfies LegacyTranslations;
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

const LoadedLocaleProvider = ({
  children,
  locale,
  onError,
}: {
  children: ReactNode;
  locale: string;
  onError?: ((error: Error) => void) | undefined;
}): JSX.Element => {
  const [messages, setMessages] = useState<FolioMessages | undefined>(() =>
    getLoadedFolioMessages(locale),
  );
  useEffect(() => {
    if (messages !== undefined) return;
    let active = true;
    void loadFolioMessages(locale).then(
      (loaded) => {
        if (active) setMessages(loaded);
        return undefined;
      },
      (cause: unknown) => {
        if (!active) return undefined;
        const error = cause instanceof Error ? cause : new Error(String(cause));
        onError?.(error);
        setMessages(english);
        return undefined;
      },
    );
    return () => {
      active = false;
    };
  }, [locale, messages, onError]);
  if (messages === undefined) return localeLoading;
  return (
    <IntlProvider locale={locale} messages={messages}>
      {children}
    </IntlProvider>
  );
};

const localeLoading = <EditorLoadingIndicator />;

export const LocaleProvider = ({
  children,
  i18n,
  locale,
  onError,
}: LocaleProviderProps): JSX.Element => {
  const resolved = resolveLocale(locale, i18n);
  return (
    <LoadedLocaleProvider key={resolved} locale={resolved} onError={onError}>
      {children}
    </LoadedLocaleProvider>
  );
};

export const useTranslation = (): { t: TFunction } => {
  const t = useTranslations("folio");
  return { t: t as unknown as TFunction };
};

export const DocxEditor: ForwardRefExoticComponent<
  DocxEditorProps & {
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
  } & RefAttributes<DocxEditorRef>
> = forwardRef<DocxEditorRef, LegacyDocxEditorProps>(
  (
    {
      agentPanel: _agentPanel,
      colorMode: _colorMode,
      commentsSidebarOpen: _commentsSidebarOpen,
      disableFindReplaceShortcuts,
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
      keyboardShortcuts,
      ...props
    },
    ref,
  ) => (
    <LocaleProvider i18n={i18n} locale={locale} onError={props.onError}>
      <FolioDocxEditor
        {...props}
        // The legacy flag only ever switched the shortcuts off; an explicit
        // scope from a migrated call site is the finer statement and wins.
        keyboardShortcuts={
          keyboardShortcuts ??
          (disableFindReplaceShortcuts === true
            ? "none"
            : DEFAULT_KEYBOARD_SHORTCUT_SCOPE)
        }
        ref={ref}
      />
    </LocaleProvider>
  ),
);

DocxEditor.displayName = "EigenpalCompatDocxEditor";

const resolveLocale = (
  locale: string | undefined,
  i18n: LegacyTranslations | undefined,
): string => locale ?? i18n?._lang ?? "en";
