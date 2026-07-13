import type { Document, SectionProperties } from "../types/document";
import type { PageHeaderFooterRefs } from "./types";

type InheritedHeaderFooterRefs = Pick<
  PageHeaderFooterRefs,
  "headerDefault" | "headerFirst" | "headerEven" | "footerDefault" | "footerFirst" | "footerEven"
>;

const resolveSectionRefs = (
  props: SectionProperties,
  inherited: InheritedHeaderFooterRefs,
  documentEvenAndOddHeaders: boolean,
): PageHeaderFooterRefs => {
  const refs: PageHeaderFooterRefs = {
    ...inherited,
    evenAndOddHeaders: props.evenAndOddHeaders ?? documentEvenAndOddHeaders,
  };
  if (props.titlePg !== undefined) {
    refs.titlePg = props.titlePg;
  }
  for (const ref of props.headerReferences ?? []) {
    if (ref.type === "default") {
      refs.headerDefault = ref.rId;
    } else if (ref.type === "first") {
      refs.headerFirst = ref.rId;
    } else {
      refs.headerEven = ref.rId;
    }
  }
  for (const ref of props.footerReferences ?? []) {
    if (ref.type === "default") {
      refs.footerDefault = ref.rId;
    } else if (ref.type === "first") {
      refs.footerFirst = ref.rId;
    } else {
      refs.footerEven = ref.rId;
    }
  }
  return refs;
};

const inheritedRefsFrom = (refs: PageHeaderFooterRefs): InheritedHeaderFooterRefs => ({
  ...(refs.headerDefault === undefined ? {} : { headerDefault: refs.headerDefault }),
  ...(refs.headerFirst === undefined ? {} : { headerFirst: refs.headerFirst }),
  ...(refs.headerEven === undefined ? {} : { headerEven: refs.headerEven }),
  ...(refs.footerDefault === undefined ? {} : { footerDefault: refs.footerDefault }),
  ...(refs.footerFirst === undefined ? {} : { footerFirst: refs.footerFirst }),
  ...(refs.footerEven === undefined ? {} : { footerEven: refs.footerEven }),
});

export const resolveSectionHeaderFooterRefs = (
  documentModel: Document | null | undefined,
): PageHeaderFooterRefs[] | undefined => {
  const body = documentModel?.package.document;
  if (!body) {
    return undefined;
  }
  const sections: SectionProperties[] = [];
  if (body.sections && body.sections.length > 0) {
    sections.push(...body.sections.map(({ properties }) => properties));
  } else if (body.finalSectionProperties) {
    sections.push(body.finalSectionProperties);
  }
  if (sections.length === 0) {
    return undefined;
  }

  const documentEvenAndOddHeaders = documentModel.package.settings?.evenAndOddHeaders === true;
  let inherited: InheritedHeaderFooterRefs = {};
  return sections.map((props) => {
    const refs = resolveSectionRefs(props, inherited, documentEvenAndOddHeaders);
    inherited = inheritedRefsFrom(refs);
    return refs;
  });
};
