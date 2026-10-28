/**
 * Hostile, malformed and unusual XML, shared by every reader's golden test.
 *
 * Each input is a whole WordprocessingML main part so that package-level
 * readers can receive it as `word/document.xml`; the oddity sits inside the
 * body, or replaces the document outright where the oddity is structural.
 */
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const body = (content: string): string =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${content}</w:body></w:document>`;

const run = (text: string): string =>
  `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

const nested = (depth: number): string =>
  `${"<w:sdt>".repeat(depth)}${"</w:sdt>".repeat(depth)}`;

export const XML_EDGE_INPUTS: ReadonlyArray<
  readonly [name: string, xml: string]
> = [
  ["ordinary", body(run("plain"))],
  ["built-in entities", body(run("&lt;&gt;&amp;&quot;&apos;"))],
  ["character references", body(run("&#65;&#x42;&#X43;&#x1F642;"))],
  ["CRLF text", body(run("a\r\nb\rc"))],
  ["CRLF and tab attribute", body('<w:p w:rsidR="a\r\nb\tc"/>')],
  ["CDATA", body(run("<![CDATA[x < y & z]]>"))],
  ["comment and PI", body(`<!-- note --><?pi data?>${run("t")}`)],
  ["mismatched end tag", body("<w:p><w:r></w:p></w:r>")],
  [
    "unclosed element",
    body("<w:p><w:r>").replace("</w:body></w:document>", ""),
  ],
  ["stray end tag", body("</w:p>")],
  ["two roots", `${body("")}<w:document xmlns:w="${W}"/>`],
  ["text before root", `junk${body("")}`],
  ["text after root", `${body("")}junk`],
  ["duplicate attribute", body('<w:p w:rsidR="1" w:rsidR="2"/>')],
  ["attribute without value", body("<w:p w:rsidR/>")],
  ["unquoted attribute", body("<w:p w:rsidR=1/>")],
  ["attributes without separator", body(`<w:p w:a="1"w:b="2"/>`)],
  ["less-than in attribute", body('<w:p w:rsidR="a<b"/>')],
  ["HTML entity in text", body(run("a&nbsp;b&copy;c&euro;"))],
  ["HTML entity in attribute", body('<w:p w:rsidR="a&nbsp;b"/>')],
  ["unknown entity", body(run("&custom;"))],
  ["bare ampersand", body(run("a & b"))],
  ["ampersand without semicolon", body(run("a &amp b"))],
  ["null character reference", body(run("a&#0;b"))],
  ["C0 character reference", body(run("a&#1;b"))],
  ["surrogate character reference", body(run("a&#xD800;b"))],
  ["out-of-range character reference", body(run("a&#1114112;b"))],
  ["lenient character reference", body(run("a&#12ab;b"))],
  ["raw control character", body(run("a\u0001b"))],
  ["CDATA terminator in text", body(run("a]]>b"))],
  ["double hyphen in comment", body("<!-- a -- b -->")],
  ["DOCTYPE", `<!DOCTYPE w:document>${body("")}`],
  [
    "DOCTYPE entity expansion",
    `<?xml version="1.0"?><!DOCTYPE w:document [<!ENTITY a "aaaaaaaaaa"><!ENTITY b "&a;&a;&a;&a;&a;&a;&a;&a;&a;&a;">]><w:document xmlns:w="${W}"><w:body>${run("&b;")}</w:body></w:document>`,
  ],
  ["prototype attribute", body('<w:p __proto__="x"/>')],
  ["constructor element", body("<constructor/>")],
  ["constructor attribute", body('<w:p constructor="x"/>')],
  ["name starting with a digit", body("<w:1p/>")],
  ["name with two colons", body("<w:a:p/>")],
  ["unbound element prefix", body("<x:p/>")],
  ["unbound attribute prefix", body('<w:p x:id="1"/>')],
  ["declaration inside the root", body('<?xml version="1.0"?>')],
  ["declaration after whitespace", ` ${body("")}`],
  ["byte-order mark", `${String.fromCodePoint(0xfeff)}${body("")}`],
  ["unclosed CDATA", body("<![CDATA[x")],
  ["unclosed comment", body("<!-- x")],
  ["depth 99", body(nested(97))],
  ["depth 101", body(nested(99))],
  ["depth 150", body(nested(148))],
  [
    "many attributes",
    body(
      `<w:p ${Array.from({ length: 5000 }, (_, index) => `w:a${String(index)}="${String(index)}"`).join(" ")}/>`,
    ),
  ],
  ["space before tag name", body("< w:p/>")],
  ["space after self-closing slash", body("<w:p/ >")],
  ["space in end tag", body("<w:p></w:p >")],
  ["empty document", ""],
  ["whitespace document", "   "],
  ["xml prefix rebound", body(`<w:p xmlns:xml="urn:other"/>`)],
  [
    "namespace URI with reference",
    body(`<w:p xmlns:x="urn:a&amp;b" x:id="1"/>`),
  ],
  [
    "numeric identifier out of range",
    body('<w:p><w:bookmarkStart w:id="99999999999" w:name="b"/></w:p>'),
  ],
  [
    "numeric identifier with reference",
    body('<w:p><w:bookmarkStart w:id="&#49;&#50;" w:name="b"/></w:p>'),
  ],
  [
    "numeric identifier with HTML entity",
    body('<w:p><w:bookmarkStart w:id="1&nbsp;" w:name="b"/></w:p>'),
  ],
  [
    "relationship attribute",
    body('<w:p><w:hyperlink r:id="rIdMissing"><w:r/></w:hyperlink></w:p>'),
  ],
];

/** Single captured elements, as property-set preservation receives them. */
export const XML_EDGE_FRAGMENTS: ReadonlyArray<
  readonly [name: string, xml: string]
> = [
  ["element", '<w:b w:val="1"/>'],
  ["element with comment after", "<w:b/><!-- c -->"],
  ["element with comment before", "<!-- c --><w:b/>"],
  ["element with PI after", "<w:b/><?pi x?>"],
  ["two elements", "<w:b/><w:i/>"],
  ["text and element", "x<w:b/>"],
  ["element and text", "<w:b/>x"],
  ["CDATA only", "<![CDATA[x]]>"],
  ["comment only", "<!-- c -->"],
  ["declaration", '<?xml version="1.0"?><w:b/>'],
  ["nested", "<w:rFonts><w:x/></w:rFonts>"],
  ["foreign namespace", '<x:y xmlns:x="urn:x"/>'],
  ["unbound prefix", "<x:y/>"],
  ["unbound attribute prefix", '<w:b x:val="1"/>'],
  ["rebinding w", '<w:b xmlns:w="urn:other"/>'],
  [
    "duplicate expanded attribute",
    '<w:b xmlns:a="urn:a" xmlns:b="urn:a" a:x="1" b:x="2"/>',
  ],
  ["duplicate attribute", '<w:b w:val="1" w:val="2"/>'],
  ["xml prefix bound elsewhere", '<w:b xmlns:xml="urn:x"/>'],
  ["xmlns prefix", '<w:b xmlns:xmlns="urn:x"/>'],
  ["empty prefix binding", '<w:b xmlns:a=""/>'],
  ["reserved xml-prefixed binding", '<w:b xmlns:xmlfoo="urn:x"/>'],
  ["HTML entity", '<w:b w:val="&nbsp;"/>'],
  ["built-in entity", '<w:b w:val="&amp;"/>'],
  ["character reference", '<w:b w:val="&#65;"/>'],
  ["invalid character reference", '<w:b w:val="&#0;"/>'],
  ["DOCTYPE", "<!DOCTYPE x><w:b/>"],
  ["ENTITY", '<!ENTITY x "y"><w:b/>'],
  ["mismatched", "<w:b><w:i></w:b></w:i>"],
  ["unclosed", "<w:b>"],
  ["prototype attribute", '<w:b __proto__="x"/>'],
  ["constructor element", "<constructor/>"],
  ["whitespace around", "  <w:b/>\n"],
  ["empty", ""],
  ["depth 101", `${"<w:x>".repeat(101)}${"</w:x>".repeat(101)}`],
];
