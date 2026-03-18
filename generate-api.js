/**
 * 論文文件產生核心（供 server.js 與 generate.js 共用）
 */
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel,
  AlignmentType, PageBreak, TableOfContents,
  Footer, PageNumber, NumberFormat, LineRuleType,
} = require('docx');

const FONT_ZH = '標楷體';
const FONT_EN = 'Times New Roman';
const SIZE_BODY = 24;
const SIZE_H1   = 28;
const SIZE_H2   = 26;
const SIZE_COVER_UNIV  = 32;
const SIZE_COVER_TITLE = 36;

const cm = (n) => Math.round(n * 567);
const isZh = (str) => /[\u4e00-\u9fff]/.test(str || '');

function para(text, opts = {}) {
  const {
    size = SIZE_BODY, bold = false, align = AlignmentType.JUSTIFIED,
    spaceBefore = 0, spaceAfter = 0, indent = 0, lineSpacing = null, font = null,
  } = opts;
  const runFont = font || (isZh(text) ? FONT_ZH : FONT_EN);
  return new Paragraph({
    alignment: align,
    spacing: {
      before: spaceBefore, after: spaceAfter,
      ...(lineSpacing ? { line: lineSpacing, lineRule: LineRuleType.AUTO } : {}),
    },
    indent: indent ? { firstLine: indent } : undefined,
    children: [new TextRun({ text, size, bold, font: { name: runFont } })],
  });
}

function mixedPara(runs, opts = {}) {
  const { align = AlignmentType.JUSTIFIED, spaceBefore = 0, spaceAfter = 0, lineSpacing = null } = opts;
  return new Paragraph({
    alignment: align,
    spacing: { before: spaceBefore, after: spaceAfter,
      ...(lineSpacing ? { line: lineSpacing, lineRule: LineRuleType.AUTO } : {}) },
    children: runs.map(r => new TextRun({
      text: r.text, size: r.size || SIZE_BODY, bold: r.bold || false,
      font: { name: r.font || (isZh(r.text) ? FONT_ZH : FONT_EN) },
    })),
  });
}

function emptyLine(n = 1) {
  return Array.from({ length: n }, () =>
    new Paragraph({ children: [new TextRun({ text: '', size: SIZE_BODY })] }));
}

function pageBreak() {
  return new Paragraph({ children: [new TextRun({ break: 1 })] });
}

function chToOrdinal(n) {
  const m = ['零','一','二','三','四','五','六','七','八','九','十'];
  if (n <= 10) return m[n];
  if (n < 20) return '十' + (n % 10 ? m[n % 10] : '');
  return m[Math.floor(n/10)] + '十' + (n % 10 ? m[n % 10] : '');
}

function chapterHeading(num, titleZh) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    alignment: AlignmentType.CENTER,
    spacing: { before: cm(0.5), after: cm(0.5) },
    children: [new TextRun({ text: `第${chToOrdinal(num)}章、${titleZh}`, size: SIZE_H1, bold: true, font: { name: FONT_ZH } })],
  });
}

function sectionHeading(number, title) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    alignment: AlignmentType.LEFT,
    spacing: { before: cm(0.3), after: cm(0.2) },
    children: [new TextRun({ text: `${number}、${title}`, size: SIZE_H2, bold: true, font: { name: FONT_ZH } })],
  });
}

function bodyParas(text) {
  if (!text) return [para('（內容待填）')];
  return text.trim().split('\n').map(line =>
    para(line.trim(), { indent: SIZE_BODY * 2, lineSpacing: 480 })
  );
}

// ── 各節 ──
function buildCoverPage(cfg) {
  return [
    ...emptyLine(2),
    para(cfg.university + cfg.college, { size: SIZE_COVER_UNIV, bold: true, align: AlignmentType.CENTER }),
    para(cfg.program, { size: SIZE_COVER_UNIV, bold: true, align: AlignmentType.CENTER, spaceAfter: cm(0.3) }),
    para(cfg.program_en, { size: SIZE_BODY, align: AlignmentType.CENTER, font: FONT_EN }),
    para(cfg.college_en, { size: SIZE_BODY, align: AlignmentType.CENTER, font: FONT_EN }),
    para(cfg.university_en, { size: SIZE_BODY, align: AlignmentType.CENTER, font: FONT_EN, spaceAfter: cm(0.5) }),
    para(cfg.thesis_type_en || 'Master Thesis', { size: SIZE_BODY, align: AlignmentType.CENTER, font: FONT_EN, spaceAfter: cm(1) }),
    para(cfg.title_zh, { size: SIZE_COVER_TITLE, bold: true, align: AlignmentType.CENTER, spaceAfter: cm(0.3) }),
    para(cfg.title_en, { size: SIZE_BODY + 2, bold: true, align: AlignmentType.CENTER, font: FONT_EN, spaceAfter: cm(1) }),
    para(cfg.author_zh, { size: SIZE_BODY, align: AlignmentType.CENTER, spaceAfter: cm(0.2) }),
    para(cfg.author_en, { size: SIZE_BODY, align: AlignmentType.CENTER, font: FONT_EN, spaceAfter: cm(1) }),
    ...(cfg.advisors || []).map(a => para(`指導教授：${a.name_zh} 博士`, { size: SIZE_BODY, align: AlignmentType.CENTER })),
    ...(cfg.advisors || []).map(a => para(`Advisor: ${a.name_en}, ${a.title_en}`, { size: SIZE_BODY, align: AlignmentType.CENTER, font: FONT_EN })),
    ...emptyLine(2),
    para(`中華民國${cfg.year_zh}年${cfg.month_zh}月`, { size: SIZE_BODY, align: AlignmentType.CENTER }),
    para(`${cfg.month_en} ${cfg.year_en}`, { size: SIZE_BODY, align: AlignmentType.CENTER, font: FONT_EN }),
  ];
}

function buildApprovalPage() {
  return [
    pageBreak(),
    para('論文審定書', { size: SIZE_H1, bold: true, align: AlignmentType.CENTER, spaceAfter: cm(0.5) }),
    para('（請黏貼學校核發之論文審定書）', { size: SIZE_BODY, align: AlignmentType.CENTER }),
  ];
}

function buildAcknowledgements(cfg) {
  const lines = (cfg.acknowledgements || '').trim().split('\n');
  return [
    pageBreak(),
    para('感言與誌謝', { size: SIZE_H1, bold: true, align: AlignmentType.CENTER, spaceAfter: cm(0.5) }),
    ...lines.map(l => para(l, { indent: SIZE_BODY * 2, lineSpacing: 480 })),
  ];
}

function buildAbstractZh(cfg) {
  const lines = (cfg.abstract_zh || '').trim().split('\n');
  const kw = (cfg.keywords_zh || []).join('、');
  return [
    pageBreak(),
    para('中文摘要', { size: SIZE_H1, bold: true, align: AlignmentType.CENTER, spaceAfter: cm(0.5) }),
    ...lines.map(l => para(l, { indent: SIZE_BODY * 2, lineSpacing: 480 })),
    ...emptyLine(1),
    mixedPara([{ text: '關鍵詞：', bold: true }, { text: kw }]),
  ];
}

function buildAbstractEn(cfg) {
  const lines = (cfg.abstract_en || '').trim().split('\n');
  const kw = (cfg.keywords_en || []).join(', ');
  return [
    pageBreak(),
    para('ABSTRACT', { size: SIZE_H1, bold: true, align: AlignmentType.CENTER, font: FONT_EN, spaceAfter: cm(0.3) }),
    mixedPara([{ text: 'NAME：', bold: true, font: FONT_EN }, { text: cfg.author_en, font: FONT_EN }]),
    mixedPara([{ text: 'MONTH/YEAR：', bold: true, font: FONT_EN }, { text: `${cfg.month_en}, ${cfg.year_en}`, font: FONT_EN }]),
    mixedPara([{ text: 'ADVISER：', bold: true, font: FONT_EN }, { text: (cfg.advisors || []).map(a => `${a.name_en}, ${a.title_en}`).join('；'), font: FONT_EN }]),
    mixedPara([{ text: 'TITLE：', bold: true, font: FONT_EN }, { text: cfg.title_en, font: FONT_EN }], { spaceAfter: cm(0.3) }),
    ...lines.map(l => para(l, { indent: SIZE_BODY * 2, lineSpacing: 480, font: FONT_EN })),
    ...emptyLine(1),
    mixedPara([{ text: 'Keywords: ', bold: true, font: FONT_EN }, { text: kw, font: FONT_EN }]),
  ];
}

function buildTOC() {
  return [
    pageBreak(),
    para('目錄', { size: SIZE_H1, bold: true, align: AlignmentType.CENTER, spaceAfter: cm(0.5) }),
    new TableOfContents('目錄', { hyperlink: true, headingStyleRange: '1-2' }),
  ];
}

function buildChapters(cfg) {
  const result = [];
  for (const ch of cfg.chapters || []) {
    result.push(pageBreak());
    result.push(chapterHeading(ch.number, ch.title_zh));
    for (const sec of ch.sections || []) {
      result.push(sectionHeading(sec.number, sec.title));
      result.push(...bodyParas(sec.content));
    }
  }
  return result;
}

function buildReferences(cfg) {
  return [
    pageBreak(),
    para('參考文獻', { size: SIZE_H1, bold: true, align: AlignmentType.CENTER, spaceAfter: cm(0.5) }),
    ...(cfg.references || []).map(ref =>
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: 480, lineRule: LineRuleType.AUTO },
        indent: { left: SIZE_BODY * 2, hanging: SIZE_BODY * 2 },
        children: [new TextRun({ text: ref, size: SIZE_BODY, font: { name: isZh(ref) ? FONT_ZH : FONT_EN } })],
      })
    ),
  ];
}

async function buildDoc(cfg) {
  const children = [
    ...buildCoverPage(cfg),
    ...buildApprovalPage(),
    ...buildAcknowledgements(cfg),
    ...buildAbstractZh(cfg),
    ...buildAbstractEn(cfg),
    ...buildTOC(),
    ...buildChapters(cfg),
    ...buildReferences(cfg),
  ];

  const doc = new Document({
    creator: cfg.author_zh,
    title: cfg.title_zh,
    styles: {
      default: {
        document: { run: { font: { name: FONT_ZH }, size: SIZE_BODY } },
      },
      paragraphStyles: [
        {
          id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal',
          run: { size: SIZE_H1, bold: true, font: { name: FONT_ZH } },
          paragraph: { alignment: AlignmentType.CENTER, spacing: { before: cm(0.5), after: cm(0.5) } },
        },
        {
          id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal',
          run: { size: SIZE_H2, bold: true, font: { name: FONT_ZH } },
          paragraph: { spacing: { before: cm(0.3), after: cm(0.2) } },
        },
      ],
    },
    sections: [{
      properties: {
        page: {
          margin: { top: cm(3), right: cm(3), bottom: cm(3), left: cm(3.5) },
          size: { width: cm(21), height: cm(29.7) },
        },
        pageNumberFormatType: NumberFormat.LOWER_ROMAN,
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ children: [PageNumber.CURRENT], font: { name: FONT_EN }, size: SIZE_BODY })],
          })],
        }),
      },
      children,
    }],
  });

  return Packer.toBuffer(doc);
}

module.exports = { buildDoc };
