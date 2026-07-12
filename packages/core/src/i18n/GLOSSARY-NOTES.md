# folio glossary — research notes (2026-07-12)

Canonical terms mimic **Microsoft Word's localized UI** (folio is a .docx editor; users
come from Word). Where Word and LibreOffice Writer diverge, Word's term is canonical and
the LibreOffice variant is recorded here, not in `forbidden`. `forbidden` lists only
well-attested wrong/nonstandard variants (Traditional-Chinese terms in zh-CN, pt-PT terms
in pt-BR, calques both suites avoid).

## Method / primary sources

Localized Microsoft Support articles use the official localized Office UI strings, so
they were used as the primary attestation source. Article IDs (same ID works per locale
as `https://support.microsoft.com/<locale>/office/<id>`):

- **Word keyboard shortcuts** `95ef89dd-7142-4b50-afb2-f762f663ceb2` — fetched for:
  cs-cz, de-de, fr-fr, es-es, pt-br, pl-pl, tr-tr, hu-hu, sk-sk, zh-cn, ar-sa, he-il,
  et-ee, lt-lt, lv-lv. (hi-in serves English only; see Hindi section.)
  Covers: save/undo/redo/copy/paste/cut/print/find/replace, bold/italic/underline,
  super/subscript, delete, footnote/endnote/page break/comment insertion, track changes,
  header/footer, hyperlink, paragraph, table, style.
- **Track changes in Word** `197ba630-0f5f-4a8e-9a77-3712475e806a` — cs-cz, zh-cn, hu-hu
  (track changes, revision, accept, reject, markup).
- **Insert footnotes and endnotes** `61f3fb1a-4717-414c-9a8f-015a5f3ff4cb` — cs-cz,
  zh-cn, sk-sk.
- **Change page orientation** `9b5ac1af-9998-4a37-962b-a82b689572a9` — cs-cz, zh-cn
  (orientation, portrait, landscape, page setup, margins).
- Web search `site:support.microsoft.com cs-cz "konec oddílu"` — confirmed Czech
  section break = "konec oddílu" ("Vložit konec oddílu", "Odstranění konce oddílu").
- LibreOffice divergence checks:
  - https://help.libreoffice.org/latest/cs/text/swriter/guide/page_break.html —
    LO cs page break = "zalomení stránky" (menu "Vložit – Další zalomení – Ruční zalomení").
  - https://help.libreoffice.org/latest/de/text/shared/guide/redlining_enter.html —
    LO de track changes = "Bearbeiten – Änderungen – Aufzeichnen" / "Änderungen verfolgen",
    "Änderungen verwalten".

Terms not covered by the fetched articles (e.g. bookmark, watermark, ruler, orientation
in most locales) come from platform knowledge of the localized Word UI / former Microsoft
Language Portal terminology; low-confidence cases were omitted rather than guessed (listed
per locale below).

## Per-locale notes

### cs (thorough; sourced from cs-cz support articles)
- Sourced: Sledování změn, revize, komentář, poznámka pod čarou, vysvětlivka, záhlaví,
  zápatí, konec stránky, konec oddílu (via search), obsah, hypertextový odkaz, orientace,
  na výšku / na šířku, Vzhled stránky, okraje, tučné, kurzíva, podtržení, přeškrtnutí,
  horní/dolní index, styl, odstavec, tabulka, Uložit, Přijmout, Odmítnout, Nahradit,
  Vyjmout, Kopírovat, Vložit, Odstranit.
- Word cs vs LibreOffice cs: page break "konec stránky" (Word) vs **"zalomení stránky"**
  (LO, sourced above); section break "konec oddílu" (Word) vs "zalomení oddílu"-style
  wording (LO). LO variants deliberately NOT forbidden.
- endnote: Word cs = "vysvětlivka"; forbidden "koncová poznámka" (literal calque that
  translators produce; neither Word nor LO cs uses it).
- insert vs paste are both "Vložit" in Czech Word (Insert tab = "Vložení"); this collision
  is faithful to the platform.
- print: Word cs backstage/command label is the nominal "Tisk" (the infinitive
  "Vytisknout" appears only in prose); glossary uses "Tisk".
- undo/redo: Word cs QAT labels "Zpět" / "Znovu" (platform knowledge; support article has
  only prose descriptions).

### de (sourced: de-de shortcuts article; LO help for divergences)
- Sourced: Änderungen nachverfolgen, Fußnote, Endnote, Seitenumbruch, Kopfzeile,
  Fußzeile, Formatvorlage, Fett, Kursiv, Unterstrichen, Hochgestellt, Tiefgestellt,
  Wiederholen (redo), Speichern, Ersetzen, Löschen, Kommentar, Hyperlink, Absatz, Tabelle.
- Word vs LO: bookmark **Textmarke** (Word) vs "Lesezeichen" (LO); style
  **Formatvorlage** (Word) vs "(Absatz)Vorlage" (LO); redo **Wiederholen** (Word) vs
  "Wiederherstellen" (LO); track changes "Änderungen nachverfolgen" (Word) vs
  "Änderungen verfolgen / Aufzeichnen" (LO, sourced).
- forbidden "Änderungen verfolgen": kept per the task's own example, even though it is
  also LO's menu wording — flagged as a debatable call.
- forbidden "Stil": common calque for style; both Word (Formatvorlage) and LO (Vorlage)
  avoid it.
- "Ausrichtung" is BOTH alignment and page orientation in German Word (genuine collision).
- Ribbon Word now often shows "Link" for hyperlink; the terminology/dialog term is still
  "Hyperlink" (kept).
- Cut: article prose showed "Schneiden"; the actual UI command is "Ausschneiden"
  (standard Windows/Office term).

### es (sourced: es-es shortcuts article)
- All verbs + formatting sourced. Word vs LO: TOC "Tabla de contenido" (Word) vs
  **"Sumario"** (LO); hyperlink "hipervínculo" (Word) vs "hiperenlace" (LO).
- forbidden "Salvar" (save): anglicism; both suites use "Guardar".
- portrait/landscape: Word es uses "Vertical" / "Horizontal" (not "retrato/paisaje").

### fr (sourced: fr-fr shortcuts article)
- Sourced: Enregistrer, Annuler, Rétablir, note de bas de page, note de fin, saut de
  page, suivi des modifications, lien hypertexte, exposant, indice, gras/italique/souligné.
- Word vs LO: bookmark **signet** (Word) vs "repère de texte" (LO); reject **Refuser**
  (Word review UI) vs "Rejeter" (LO).
- forbidden "Sauvegarder" (save): MS French reserves it for backup; UI verb is
  "Enregistrer". forbidden "indentation" (indent): anglicism; both suites use "retrait".

### pt-BR (sourced: pt-br shortcuts article)
- Sourced: Salvar, Desfazer, Refazer, Recortar, Excluir, Localizar, Substituir,
  Controle de alterações, nota de rodapé, nota de fim, quebra de página, cabeçalho,
  rodapé, hiperlink, sobrescrito, subscrito.
- TOC: Word pt-BR = **"Sumário"** (LO pt-BR also "Sumário"); forbidden
  "Tabela de Conteúdo" (frequent bad calque). "Índice" NOT forbidden because it is a real
  Word pt-BR feature (back-of-book index) and the pt-PT TOC term; flagged as debatable.
- bookmark: Word pt-BR = **"Indicador"**; LO pt-BR uses "Marcador"/"Marca-página" (notes
  only).
- forbidden: "Guardar" (pt-PT save), "Deletar" (colloquial anglicism), "Cortar" (pt-PT
  cut; pt-BR Word = Recortar), "indentação" (code anglicism; Word/LO use recuo).
- Track Changes: feature noun "Controle de Alterações" (article); the ribbon toggle reads
  "Controlar Alterações" — verb form, same term family.
- watermark: Word pt-BR ribbon spells "Marca-d'água" (hyphenated); "marca d'água" also
  circulates.

### pl (sourced: pl-pl shortcuts article)
- Sourced: Śledzenie zmian, przypis dolny, przypis końcowy, podział strony, nagłówek,
  stopka, hiperłącze, pogrubienie, kursywa, podkreślenie, indeks górny/dolny, akapit,
  Zapisz-family verbs.
- forbidden "paragraf" (paragraph): false friend — in Polish it means a statutory
  section; Word/LO use "akapit".
- findAndReplace: Word pl dialog title is "Znajdowanie i zamienianie" (gerund style);
  LO pl uses "Znajdź i zamień" (notes only). Debatable call, see below.
- redo: Word pl = "Wykonaj ponownie"; LO pl = "Ponów" (notes only).
- revision: omitted (Word pl markup vocabulary uses "poprawki/znaczniki" inconsistently;
  no clean single noun attested).

### tr (sourced: tr-tr shortcuts article)
- Sourced: Kaydet, Geri Al, Yinele, değişiklik izleme, dipnot, son not, sayfa sonu,
  açıklama, üstbilgi/altbilgi (article spells "Üst bilgi"/"Alt bilgi"; ribbon uses the
  closed compounds "Üstbilgi"/"Altbilgi" — closed form kept), köprü, üst/alt simge.
- comment: Word tr = **"açıklama"**; forbidden "yorum" (generic term used by
  Google Docs and casual translations; both Word and LO tr use açıklama).
- underline: canonical "Altı çizili" (font-dialog/tooltip form); the shortcuts article
  also has the noun "alt çizgi".
- hyperlink: classic Word tr term "Köprü" kept; newer ribbons sometimes show "Bağlantı"
  (notes only).

### hu (sourced: hu-hu shortcuts + track-changes articles)
- Sourced: Változások követése (feature), **korrektúra = revision/markup**, Elfogadás,
  Elvetés, lábjegyzet, végjegyzet, oldaltörés, megjegyzés, élőfej, élőláb, hivatkozás,
  félkövér, dőlt, aláhúzott, felső/alsó index, bekezdés, Mentés-family verbs.
- Desktop Word hu labels the Track Changes ribbon toggle "Korrektúra"; the support
  article and feature name use "Változások követése". Canonical = "Változások követése",
  with korrektúra assigned to `revision`. Debatable call, see below.
- forbidden "fejléc"/"lábléc": generic IT header/footer words (email, tables, web);
  Word and LO hu both use élőfej/élőláb for page header/footer.
- Hungarian Word uses nominal command labels (Mentés, Kivágás, Beszúrás, Nyomtatás);
  glossary follows that convention.
- **redo omitted**: article had only prose ("Hajtsa végre újra"); Word hu QAT label
  (Ismét/Ismétlés?) vs LO hu "Újra" could not be attested confidently.

### sk (sourced: sk-sk shortcuts + footnotes articles)
- Sourced: Sledovanie zmien (article verb form "Sledovať zmeny"), poznámka pod čiarou,
  **vysvetlivka** (endnote — NOT "koncová poznámka"), zlom strany, hlavička, päta,
  hypertextové prepojenie, tučné, kurzíva, podčiarknutie, horný/dolný index, odsek,
  Uložiť, Prilepiť (paste — distinctive vs cs "Vložit"), Vystrihnúť, Zopakovať, Tlačiť.
- Word vs LO: watermark **Vodotlač** (Word sk) vs "vodoznak" (LO sk; notes only);
  redo "Zopakovať" (article) vs "Znova" (LO).
- undo: article prose gave "Zrušiť/Vrátiť"; actual Word sk QAT label is "Späť"
  (platform knowledge, mirrors cs "Zpět").
- Moderate-confidence (platform knowledge, no article): zlom sekcie, odsadenie,
  Nastavenie strany, Hľadať a nahradiť, prečiarknutie.

### zh-CN (sourced: zh-cn shortcuts, track-changes, footnotes, orientation articles)
- Sourced: 修订/跟踪修订 (see below), 批注, 脚注, 尾注, 页眉, 页脚, 分页符, 超链接, 加粗,
  倾斜, 下划线, 上标, 下标, 样式, 段落, 表格, 页边距, 页面设置, 纵向, 横向, 保存, 撤消,
  重做, 复制, 粘贴, 剪切, 打印, 查找, 替换, 删除, 插入, 接受, 拒绝.
- trackChanges canonical **修订** (the Review-tab ribbon toggle 审阅 > 修订); article prose
  also uses 跟踪修订 / 跟踪更改. Forbidden 追蹤修訂 (zh-TW Word).
- orientation: canonical **纸张方向** (desktop Word Layout tab); the orientation article
  (Word for the web) shows plain 方向.
- undo: MS officially spells **撤消**; the common variant 撤销 (used by LO and much other
  software) is a notes-level variant, NOT forbidden. Forbidden 復原 (zh-TW Word undo).
- redo: shortcuts article = 重做 (matches LO); desktop Word Ctrl+Y sometimes labeled 恢复
  (repeat). Canonical 重做.
- bold/italic: Word zh-CN = **加粗 / 倾斜**; LO zh-CN = 粗体 / 斜体 (notes only, not
  forbidden).
- Forbidden entries are verified zh-TW Word terms where Taiwan uses a *different word*,
  not merely traditional script: 註解 (comment), 註腳 (footnote), 章節附註 (endnote),
  頁首/頁尾 (header/footer), 分頁符號/分節符號 (breaks), 超連結 (hyperlink), 浮水印
  (watermark), 縮排 (indent), 底線 (underline), 儲存格 (cell), 版面設定 (page setup),
  尋找及取代 (find and replace), 儲存 (save), 復原 (undo), 取代 (replace), 剪下 (cut),
  貼上 (paste), 列印 (print), plus simplified generic 评论 (comment; Google Docs term).
  Pure same-word traditional spellings (e.g. 刪除, 樣式) were NOT listed as forbidden.

### ar (sourced: ar-sa shortcuts article)
- Sourced: تعقب التغييرات, حاشية سفلية, **تعليق ختامي** (endnote), فاصل صفحات, رأس الصفحة,
  تذييل الصفحة, ارتباط تشعبي, غامق, مائل, تسطير, نمط, فقرة, جدول, verbs نسخ/لصق/قص/طباعة/
  حذف/استبدال/إدراج (via "إدراج حاشية").
- superscript/subscript: article prose showed odd renderings ("رمز علوي"/"لاحقة سفلية");
  canonical uses the Word ar font-dialog terms مرتفع / منخفض (platform knowledge).
- Platform knowledge (no article): فاصل مقطعي, جدول المحتويات, إشارة مرجعية, علامة مائية,
  هامش, مسافة بادئة, محاذاة, يتوسطه خط, مسطرة, اتجاه/عمودي/أفقي, إعداد الصفحة,
  بحث واستبدال, قبول/رفض, تراجع/إعادة, حفظ.

### he (sourced: he-il shortcuts article)
- Sourced: מעקב אחר שינויים, הערת שוליים, הערת סיום, מעבר עמוד, הערה (comment),
  היפר-קישור, מודגש, נטוי, קו תחתון, סגנון, טבלה, שמור/העתק/הדבק/גזור/הדפס/החלף/בצע שוב.
- Platform knowledge: כותרת עליונה/תחתונה (header/footer — absent from article but
  standard Word he), כתב עילי/תחתי, מעבר מקטע, תוכן עניינים, סימניה, סימן מים, שוליים,
  כניסה (indent), יישור, קו חוצה, סרגל, לאורך/לרוחב, הגדרת עמוד, חיפוש והחלפה, קבל/דחה,
  הוסף, מחק, בטל.
- paragraph: MS he traditionally spells פיסקה (kept); the plain-spelling variant פסקה is
  common elsewhere — spelling variant only, not forbidden.
- orientation: plain כיוון kept (Word he layout wording; sometimes כיוון הדפסה).

### et (sourced: et-ee shortcuts article; deliberately sparse)
- Sourced: allmärkus, lõpumärkus, leheküljepiir, kommentaar, muutuste jälitus, päis,
  jalus (article rendered the inflected/typo "jalas"; nominative is jalus), hüperlink,
  lõik, tabel, **laad** (style — distinctive; LO et uses "stiil"), paks, kursiiv,
  allakriipsutus, üla-/allindeks, salvesta/kopeeri/kleebi/lõika/prindi/kustuta/asenda/lisa,
  tee uuesti (redo).
- undo: article prose "tühistab"; Estonian Office uses **"Võta tagasi"** for Undo
  ("Tühista" = Cancel) — platform knowledge, kept.
- Platform knowledge: sisukord, järjehoidja, vesimärk, veeris, taane, joondus, joonlaud,
  lahter/rida/veerg.
- Omitted (no confident attestation): revision, section break, strikethrough, orientation/
  portrait/landscape, page setup, find-and-replace, accept, reject.

### lt (sourced: lt-lt shortcuts article; deliberately sparse)
- Sourced: **puslapio išnaša** (footnote) / **dokumento išnaša** (endnote) — Word lt
  qualifies both, plain "išnaša" is the generic word; puslapio lūžis, komentaras,
  keitimų sekimas (article verb "Sekti keitimus"), antraštė, poraštė, hipersaitas,
  pastraipa, lentelė, stilius, paryškintasis, pasvirasis, pabrauktasis,
  viršutinis/apatinis indeksas, Įrašyti, Anuliuoti, Perdaryti, Keisti (replace),
  Naikinti, Iškirpti, Kopijuoti, Įklijuoti, Spausdinti, Įterpti.
- save: MS lt uses **Įrašyti**; the generic "Išsaugoti" (common in other software) is a
  notes-level variant, not forbidden.
- Platform knowledge: turinys, žymelė, paraštė, įtrauka, lygiuotė, liniuotė, vandens
  ženklas, langelis/eilutė/stulpelis.
- Omitted: revision, section break, strikethrough, orientation set, page setup,
  find-and-replace, accept, reject.

### lv (sourced: lv-lv shortcuts article; deliberately sparse)
- Sourced: **izmaiņu reģistrēšana** (Track Changes — distinctive Word lv term),
  lappuses pārtraukums, komentārs, galvene, kājene, hipersaite, rindkopa, tabula, stils,
  treknraksts, slīpraksts, pasvītrojums, augšraksts, apakšraksts, Saglabāt, Atsaukt,
  **Atcelt atsaukšanu** (redo), Aizstāt, Dzēst, Izgriezt, Kopēt, Ielīmēt, Drukāt, Ievietot.
- Platform knowledge: vēre / beigu vēre (footnote/endnote — MS lv terminology),
  satura rādītājs, grāmatzīme, ūdenszīme, atkāpe, līdzinājums, šūna/rinda/kolonna.
- Omitted: margin (piemale/mala not confidently attested), ruler, revision, section
  break, strikethrough, orientation set, page setup, find-and-replace, accept, reject.

### hi (deliberately very sparse)
- support.microsoft.com hi-in serves English-only content (verified on two articles), so
  no article attestation was possible. Only high-confidence Microsoft Hindi LIP/Office
  terms are included: टिप्पणी (comment), तालिका (table), अनुच्छेद (paragraph), शैली
  (style), पंक्ति (row), स्तंभ (column), बोल्ड, इटैलिक, सहेजें (save), पूर्ववत करें
  (undo), हटाएँ (delete), काटें (cut), कॉपी करें / पेस्ट करें / प्रिंट करें (modern MS
  Hindi favors transliteration for clipboard/print).
- Everything else omitted. Note for translators: modern Office Hindi increasingly
  transliterates (इन्सर्ट, फ़ॉन्ट); older LIP used Sanskritized forms — the glossary
  should be revisited if hi becomes a priority.

## Word-vs-LibreOffice divergence summary (LO variant NOT forbidden)

| Locale | Concept | Word (canonical) | LibreOffice |
|---|---|---|---|
| cs | page break | Konec stránky | Zalomení stránky (sourced) |
| de | track changes | Änderungen nachverfolgen | Änderungen (verfolgen)/Aufzeichnen (sourced) |
| de | bookmark | Textmarke | Lesezeichen |
| de | style | Formatvorlage | (Absatz)Vorlage |
| de | redo | Wiederholen | Wiederherstellen |
| es | TOC | Tabla de contenido | Sumario |
| es | hyperlink | Hipervínculo | Hiperenlace |
| fr | bookmark | Signet | Repère de texte |
| fr | reject | Refuser | Rejeter |
| pt-BR | bookmark | Indicador | Marcador / Marca-página |
| pl | find & replace | Znajdowanie i zamienianie | Znajdź i zamień |
| pl | redo | Wykonaj ponownie | Ponów |
| sk | watermark | Vodotlač | Vodoznak |
| sk | redo | Zopakovať | Znova |
| zh-CN | bold / italic | 加粗 / 倾斜 | 粗体 / 斜体 |
| zh-CN | undo spelling | 撤消 | 撤销 |
| et | style | Laad | Stiil |

## Coverage summary

- Locales filled: 16. Nouns: 34 entries; verbs: 12 entries (46 total).
- Translations filled: 669 of 736 possible cells. Full coverage (46/46) for ar, cs, de,
  es, fr, pt-BR, sk, zh-CN; near-full for he, hu, pl, tr (45); deliberately sparse for
  et (36), lt (36), lv (34), hi (15).
- Forbidden variants: 37 total (zh-CN 20, pt-BR 5, de 2, fr 2, hu 2, cs 1, es 1, pl 1,
  tr 2 — counting shared-entry seconds like zh-CN "註解"+"评论").

## Five most debatable calls

1. **de forbidden "Änderungen verfolgen"** — it is simultaneously the task's own example
   of a forbidden variant and LibreOffice's actual menu wording. Kept forbidden (per the
   example), but if the lint should never flag LO terms, move it to notes.
2. **zh-CN trackChanges = 修订** — ribbon toggle term; Microsoft's own articles freely use
   跟踪修订/跟踪更改 in prose, and `revision` maps to the same word 修订 (a real collision
   in Chinese Word).
3. **hu trackChanges = "Változások követése", revision = "Korrektúra"** — desktop Word hu
   labels the ribbon toggle "Korrektúra"; the split follows the support article. hu redo
   was omitted entirely for lack of attestation.
4. **pl findAndReplace = "Znajdowanie i zamienianie"** — faithful to the Word pl dialog
   title, but clunky as a glossary term; LO's "Znajdź i zamień" is what most Polish users
   would type. Chosen per the "prefer Word" rule.
5. **pt-BR TOC = "Sumário" with only "Tabela de Conteúdo" forbidden** — "Índice" was NOT
   forbidden despite being a frequent mistranslation, because it is the legitimate Word
   pt-BR term for a back-of-book index (and pt-PT's TOC term); forbidding it would flag
   correct index-feature strings.

Nominal-vs-infinitive command style is preserved per platform (hu and cs/sk partly use
nominal labels: Mentés, Tisk); translators should follow the glossary form, not conjugate.
