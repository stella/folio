// Each spelling below reads or writes a painted position outside the owner.
declare const element: HTMLElement;
declare const root: ParentNode;
declare const position: number;

export const start = Number(element.dataset["pmStart"]);
export const end = element.dataset.pmEnd;
export const table = element.dataset["tablePmStart"];
element.dataset["sdtPmPos"] = "3";
export const attribute = element.getAttribute("data-pm-start");
export const spans = root.querySelectorAll("span[data-pm-start][data-pm-end]");
export const anchor = root.querySelector(`[data-pm-start="${String(position)}"]`);
export const handle = root.querySelector(".handle[data-table-pm-start]");
declare const editing: boolean;
element.dataset[editing ? "pmStart" : "decorationStart"] = "1";
export const { pmEnd: destructuredEnd } = element.dataset;
export const KEYS = { pmPos: "sdtPmPos" } as const;
