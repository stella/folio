import { describe, expect, test } from "bun:test";

import { cleanPastedHtml } from "./pasteCleanup";

describe("cleanPastedHtml — Office cruft removal", () => {
  test("strips mso-* declarations but keeps real CSS", () => {
    const out = cleanPastedHtml(
      '<p style="mso-margin-top-alt:auto; color:#FF0000; mso-pagination:widow-orphan; font-weight:bold">Hi</p>',
    );
    expect(out).not.toContain("mso-");
    expect(out).toContain("color:#FF0000");
    expect(out).toContain("font-weight:bold");
    expect(out).toContain(">Hi</p>");
  });

  test("removes a style attribute that held only mso declarations", () => {
    const out = cleanPastedHtml('<span style="mso-spacerun:yes">text</span>');
    expect(out).toBe("<span>text</span>");
  });

  test("does not corrupt a declaration whose value contains semicolons", () => {
    const out = cleanPastedHtml(
      `<div style="mso-foo:bar; background:url('data:image/png;base64,AAAB');color:red">x</div>`,
    );
    expect(out).not.toContain("mso-foo");
    expect(out).toContain("data:image/png;base64,AAAB");
    expect(out).toContain("color:red");
    expect(out).toContain(">x</div>");
  });

  test("keeps a quoted font name that uses the other quote character", () => {
    const out = cleanPastedHtml(
      `<span style="mso-bidi-font-family:'Times New Roman'; font-family:'Times New Roman'">x</span>`,
    );
    expect(out).not.toContain("mso-bidi");
    expect(out).toContain("font-family:'Times New Roman'");
    expect(out).toContain(">x</span>");
  });

  test("keeps producer class tokens so the style inliner can match them", () => {
    // The inliner resolves `.MsoNormal { ... }` rules against these classes; the
    // schema parser ignores class names, so they are harmless after parsing.
    expect(cleanPastedHtml('<p class="MsoNormal">a</p>')).toBe('<p class="MsoNormal">a</p>');
    expect(cleanPastedHtml('<p class="MsoListParagraph highlighted">a</p>')).toBe(
      '<p class="MsoListParagraph highlighted">a</p>',
    );
  });

  test("preserves both the <style> rule and its class so the inliner can apply it", () => {
    // Regression: cleanup runs before the style inliner, so it must keep the
    // MsoNormal class AND the stylesheet, or the class rule's formatting is lost.
    const out = cleanPastedHtml(
      '<style>.MsoNormal{font-size:18pt}</style><p class="MsoNormal">x</p>',
    );
    expect(out).toContain("<style>.MsoNormal{font-size:18pt}</style>");
    expect(out).toContain('<p class="MsoNormal">x</p>');
  });

  test("removes namespaced Office tags and smart tags but keeps their text", () => {
    const out = cleanPastedHtml(
      "<p>See <st1:place><st1:City>Berlin</st1:City></st1:place> today<o:p></o:p></p>",
    );
    expect(out).not.toContain("st1:");
    expect(out).not.toContain("o:p");
    expect(out).toContain("Berlin");
    expect(out).toContain("today");
  });

  test("removes a namespaced tag whose attribute value contains a quoted '>'", () => {
    expect(cleanPastedHtml('<o:p data-x="a>b">keep</o:p>')).toBe("keep");
  });

  test("removes an empty span whose attribute value contains a quoted '>'", () => {
    expect(cleanPastedHtml('<span title="x>y"></span>keep')).toBe("keep");
  });

  test("keeps a stylesheet hidden inside comment delimiters inside <style>", () => {
    const out = cleanPastedHtml('<style><!-- .keep{color:red} --></style><p class="keep">x</p>');
    // The style block (and its class rule) must survive the comment strip so the
    // downstream inliner can still resolve `.keep`.
    expect(out).toContain("<style>");
    expect(out).toContain(".keep{color:red}");
    expect(out).toContain('<p class="keep">x</p>');
  });

  test("still strips comments that sit outside a <style> block", () => {
    const out = cleanPastedHtml(
      "<!-- drop me --><style>.a{color:red}</style><!-- and me --><p>x</p>",
    );
    expect(out).not.toContain("drop me");
    expect(out).not.toContain("and me");
    expect(out).toContain("<style>.a{color:red}</style>");
    expect(out).toContain("<p>x</p>");
  });

  test("removes conditional comments including the Office xml island", () => {
    const html = "<!--[if gte mso 9]><xml><o:OfficeDocumentSettings/></xml><![endif]--><p>Body</p>";
    const out = cleanPastedHtml(html);
    expect(out).not.toContain("mso");
    expect(out).not.toContain("<xml");
    expect(out).not.toContain("OfficeDocumentSettings");
    expect(out).toBe("<p>Body</p>");
  });

  test("strips XML processing instructions, meta/link, and font tags (unwrapping content)", () => {
    const out = cleanPastedHtml(
      '<?xml version="1.0"?><meta charset="utf-8"><link rel="x"><font face="Arial">Word</font>',
    );
    expect(out).not.toContain("<?xml");
    expect(out).not.toContain("<meta");
    expect(out).not.toContain("<link");
    expect(out).not.toContain("<font");
    expect(out).toContain("Word");
  });

  test("removes empty spans, including nested ones, but keeps spans with content", () => {
    expect(cleanPastedHtml("<p><span></span><span>keep</span></p>")).toBe(
      "<p><span>keep</span></p>",
    );
    expect(cleanPastedHtml("<p><span><span></span></span>x</p>")).toBe("<p>x</p>");
  });

  test("keeps whitespace-only spans so adjacent words never merge", () => {
    const out = cleanPastedHtml("word<span> </span>word");
    expect(out).toContain("word<span> </span>word");
  });

  test("preserves a pasted table structure while stripping cell cruft", () => {
    const html =
      '<table class="MsoTableGrid"><tr><td style="mso-border-alt:solid; padding:5pt"><p class="MsoNormal">Cell</p></td></tr></table>';
    const out = cleanPastedHtml(html);
    expect(out).toContain("<table");
    expect(out).toContain("<tr>");
    expect(out).toContain("<td");
    expect(out).toContain("padding:5pt");
    expect(out).not.toContain("mso-");
    // Producer classes stay on the elements for the inliner to match.
    expect(out).toContain('class="MsoTableGrid"');
    expect(out).toContain('class="MsoNormal"');
    expect(out).toContain("Cell");
  });

  test("preserves a pasted Word list as ordered paragraphs", () => {
    const html =
      '<p class="MsoListParagraph" style="mso-list:l0 level1 lfo1; margin-left:36pt">First</p>' +
      '<p class="MsoListParagraph" style="mso-list:l0 level1 lfo1; margin-left:36pt">Second</p>';
    const out = cleanPastedHtml(html);
    expect(out).not.toContain("mso-list");
    // The list class is kept so the inliner can resolve its `<style>` rule.
    expect(out).toContain('class="MsoListParagraph"');
    expect(out).toContain("margin-left:36pt");
    expect(out).toContain("First");
    expect(out).toContain("Second");
  });
});

describe("cleanPastedHtml — safety and robustness", () => {
  test("leaves non-Office HTML essentially untouched", () => {
    const html = "<p>Plain <strong>bold</strong> text</p>";
    expect(cleanPastedHtml(html)).toBe(html);
  });

  test("does not rewrite typographic characters or non-Latin scripts", () => {
    const html = "<p>“quoted” — مرحبا</p>";
    expect(cleanPastedHtml(html)).toBe(html);
  });

  test("drops an unterminated comment without leaving a stray opener", () => {
    expect(cleanPastedHtml("safe<!--dangling")).toBe("safe");
  });

  test("empty input returns empty", () => {
    expect(cleanPastedHtml("")).toBe("");
  });

  test("handles a hostile run of namespace openers in linear time", () => {
    const evil = "<o:p>".repeat(100_000);
    const start = performance.now();
    const out = cleanPastedHtml(evil);
    expect(performance.now() - start).toBeLessThan(2000);
    expect(out).toBe("");
  });
});

describe("cleanPastedHtml — text-box anchor hijack via paste", () => {
  test("strips a data-docx-textbox-anchor marker from external HTML", () => {
    // A page outside folio could plant this marker with an id chosen to
    // collide with a real anchor elsewhere in the document; stripping it
    // makes the pasted span parse as an inert, unrecognized span instead of
    // a textBoxAnchor node.
    const html = '<p><span data-docx-textbox-anchor="0:0">​</span>Hijack</p>';
    const out = cleanPastedHtml(html);
    expect(out).not.toContain("data-docx-textbox-anchor");
    expect(out).toContain("Hijack");
  });

  test("keeps a data-docx-textbox-anchor marker inside a ProseMirror clipboard slice", () => {
    // Internal copy/paste of a real text box serializes through
    // prosemirror-view's clipboard wrapper, which stamps `data-pm-slice` on
    // the outer element. That must not be treated as external DOM.
    const html =
      '<div data-pm-slice="0 0 []"><p><span data-docx-textbox-anchor="0:0">​</span>Real box</p></div>';
    const out = cleanPastedHtml(html);
    expect(out).toContain("data-docx-textbox-anchor");
    expect(out).toContain("Real box");
  });

  test("strips a single- or double-quoted anchor attribute value", () => {
    expect(cleanPastedHtml("<span data-docx-textbox-anchor='0:0'>x</span>")).not.toContain(
      "data-docx-textbox-anchor",
    );
    expect(cleanPastedHtml('<span data-docx-textbox-anchor="0:0">x</span>')).not.toContain(
      "data-docx-textbox-anchor",
    );
  });
});
