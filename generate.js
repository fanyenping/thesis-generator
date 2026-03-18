#!/usr/bin/env node
/**
 * 台灣碩士論文文件產生器
 * 依照台灣大學 EiMBA 論文格式
 * 用法：node generate.js [設定檔路徑]
 */

const fs = require('fs');
const path = require('path');
const yaml = require('js-yaml');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel,
  AlignmentType, PageBreak, TableOfContents,
  Header, Footer, PageNumber, NumberFormat,
  SectionType, LevelFormat, convertInchesToTwip,
  BorderStyle, Table, TableRow, TableCell, WidthType,
  UnderlineType, LineRuleType,
} = require('docx');

// ── 單位換算 ──────────────────────────────────────────────
// 1 inch = 1440 twips；1 cm ≈ 567 twips
const cm = (n) => Math.round(n * 567);

// ── 讀取設定檔 ────────────────────────────────────────────
const configPath = process.argv[2] || path.join(__dirname, 'thesis_config.yaml');
if (!fs.existsSync(configPath)) {
  console.error(`找不到設定檔：${configPath}`);
  process.exit(1);
}
const cfg = yaml.load(fs.readFileSync(configPath, 'utf8'));
console.log(`✓ 讀取設定檔：${configPath}`);

// ── 字型與尺寸常數 ─────────────────────────────────────────
const FONT_ZH = '標楷體';
const FONT_EN = 'Times New Roman';
const SIZE_BODY = 24;       // 12pt (half-pt)
const SIZE_H1   = 28;       // 14pt
const SIZE_H2   = 26;       // 13pt
const SIZE_COVER_UNIV = 32; // 16pt
const SIZE_COVER_TITLE = 36;// 18pt

// ── 輔助函式 ──────────────────────────────────────────────

/** 建立純文字段落 */
function para(text, opts = {}) {
  const {
    size = SIZE_BODY, bold = false, align = AlignmentType.JUSTIFIED,
    spaceBefore = 0, spaceAfter = 0, indent = 0, lineSpacing = null,
    font = null,
  } = opts;
  const runFont = font || (isZh(text) ? FONT_ZH : FONT_EN);
  return new Paragraph({
    alignment: align,
    spacing: {
      before: spaceBefore,
      after: spaceAfter,
      ...(lineSpacing ? { line: lineSpacing, lineRule: LineRuleType.AUTO } : {}),
    },
    indent: indent ? { firstLine: indent } : undefined,
    children: [
      new TextRun({
        text,
        size,
        bold,
        font: { name: runFont },
      }),
    ],
  });
}

/** 混合中英文段落 */
function mixedPara(runs, opts = {}) {
  const {
    align = AlignmentType.JUSTIFIED,
    spaceBefore = 0, spaceAfter = 0, indent = 0, lineSpacing = null,
  } = opts;
  return new Paragraph({
    alignment: align,
    spacing: {
      before: spaceBefore,
      after: spaceAfter,
      ...(lineSpacing ? { line: lineSpacing, lineRule: LineRuleType.AUTO } : {}),
    },
    indent: indent ? { firstLine: indent } : undefined,
    children: runs.map(r => new TextRun({
      text: r.text,
      size: r.size || SIZE_BODY,
      bold: r.bold || false,
      font: { name: r.font || (isZh(r.text) ? FONT_ZH : FONT_EN) },
    })),
  });
}

/** 判斷是否含中文 */
function isZh(str) {
  return /[\u4e00-\u9fff]/.test(str || '');
}

/** 空白段落 */
function emptyLine(count = 1) {
  return Array.from({ length: count }, () =>
    new Paragraph({ children: [new TextRun({ text: '', size: SIZE_BODY })] })
  );
}

/** 分頁 */
function pageBreak() {
  return new Paragraph({
    children: [new TextRun({ break: 1 })],
  });
}

/** 章節標題（第X章） */
function chapterHeading(num, titleZh) {
  const text = `第${chToOrdinal(num)}章、${titleZh}`;
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    alignment: AlignmentType.CENTER,
    spacing: { before: cm(0.5), after: cm(0.5) },
    children: [
      new TextRun({
        text,
        size: SIZE_H1,
        bold: true,
        font: { name: FONT_ZH },
      }),
    ],
  });
}

/** 小節標題（X.X） */
function sectionHeading(number, title) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    alignment: AlignmentType.LEFT,
    spacing: { before: cm(0.3), after: cm(0.2) },
    children: [
      new TextRun({
        text: `${number}、${title}`,
        size: SIZE_H2,
        bold: true,
        font: { name: FONT_ZH },
      }),
    ],
  });
}

/** 內文段落（支援多行） */
function bodyParas(text) {
  if (!text) return [para('（內容待填）', { align: AlignmentType.JUSTIFIED })];
  return text.trim().split('\n').map(line =>
    para(line.trim(), {
      indent: SIZE_BODY * 2, // 首行縮排 2 字元
      lineSpacing: 480,       // 雙倍行距
      spaceBefore: 0,
      spaceAfter: 0,
    })
  );
}

/** 數字轉中文序數 */
function chToOrdinal(n) {
  const map = ['零','一','二','三','四','五','六','七','八','九','十'];
  if (n <= 10) return map[n];
  if (n < 20) return '十' + (n % 10 ? map[n % 10] : '');
  return map[Math.floor(n / 10)] + '十' + (n % 10 ? map[n % 10] : '');
}

// ─────────────────────────────────────────────────────────
// 各節內容建立函式
// ─────────────────────────────────────────────────────────

/** 封面頁 */
function buildCoverPage() {
  const advisorZh = cfg.advisors.map(a => `指導教授：${a.name_zh} 博士`).join('　　');
  const advisorEn = cfg.advisors.map(a => `Advisor: ${a.name_en}, ${a.title_en}`).join('\n');

  return [
    emptyLine(2),
    // 學校中文
    para(cfg.university + cfg.college, {
      size: SIZE_COVER_UNIV, bold: true,
      align: AlignmentType.CENTER, spaceAfter: 0,
    }),
    para(cfg.program, {
      size: SIZE_COVER_UNIV, bold: true,
      align: AlignmentType.CENTER, spaceAfter: cm(0.3),
    }),
    // 學校英文
    para(cfg.program_en, {
      size: SIZE_BODY, align: AlignmentType.CENTER, font: FONT_EN,
    }),
    para(cfg.college_en, {
      size: SIZE_BODY, align: AlignmentType.CENTER, font: FONT_EN,
    }),
    para(cfg.university_en, {
      size: SIZE_BODY, align: AlignmentType.CENTER, font: FONT_EN, spaceAfter: cm(0.5),
    }),
    para(cfg.thesis_type_en, {
      size: SIZE_BODY, align: AlignmentType.CENTER, font: FONT_EN, spaceAfter: cm(1),
    }),
    // 論文題目
    para(cfg.title_zh, {
      size: SIZE_COVER_TITLE, bold: true,
      align: AlignmentType.CENTER, spaceAfter: cm(0.3),
    }),
    para(cfg.title_en, {
      size: SIZE_BODY + 2, bold: true,
      align: AlignmentType.CENTER, font: FONT_EN, spaceAfter: cm(1),
    }),
    // 作者
    para(cfg.author_zh, {
      size: SIZE_BODY, align: AlignmentType.CENTER, spaceAfter: cm(0.2),
    }),
    para(cfg.author_en, {
      size: SIZE_BODY, align: AlignmentType.CENTER, font: FONT_EN, spaceAfter: cm(1),
    }),
    // 指導教授
    para(advisorZh, {
      size: SIZE_BODY, align: AlignmentType.CENTER, spaceAfter: 0,
    }),
    ...cfg.advisors.map(a =>
      para(`Advisor: ${a.name_en}, ${a.title_en}`, {
        size: SIZE_BODY, align: AlignmentType.CENTER, font: FONT_EN,
      })
    ),
    ...emptyLine(2),
    // 日期
    para(`中華民國${cfg.year_zh}年${cfg.month_zh}月`, {
      size: SIZE_BODY, align: AlignmentType.CENTER,
    }),
    para(`${cfg.month_en} ${cfg.year_en}`, {
      size: SIZE_BODY, align: AlignmentType.CENTER, font: FONT_EN,
    }),
  ].flat();
}

/** 論文審定書（佔位） */
function buildApprovalPage() {
  return [
    pageBreak(),
    para('論文審定書', {
      size: SIZE_H1, bold: true,
      align: AlignmentType.CENTER, spaceAfter: cm(0.5),
    }),
    para('（請黏貼學校核發之論文審定書）', {
      size: SIZE_BODY, align: AlignmentType.CENTER,
    }),
  ];
}

/** 誌謝 */
function buildAcknowledgements() {
  const lines = (cfg.acknowledgements || '').trim().split('\n');
  return [
    pageBreak(),
    para('感言與誌謝', {
      size: SIZE_H1, bold: true,
      align: AlignmentType.CENTER, spaceAfter: cm(0.5),
    }),
    ...lines.map(line => para(line, {
      indent: SIZE_BODY * 2,
      lineSpacing: 480,
    })),
  ];
}

/** 中文摘要 */
function buildAbstractZh() {
  const lines = (cfg.abstract_zh || '').trim().split('\n');
  const kw = (cfg.keywords_zh || []).join('、');
  return [
    pageBreak(),
    para('中文摘要', {
      size: SIZE_H1, bold: true,
      align: AlignmentType.CENTER, spaceAfter: cm(0.5),
    }),
    ...lines.map(line => para(line, {
      indent: SIZE_BODY * 2,
      lineSpacing: 480,
    })),
    ...emptyLine(1),
    mixedPara([
      { text: '關鍵詞：', bold: true },
      { text: kw },
    ]),
  ];
}

/** 英文摘要 */
function buildAbstractEn() {
  const lines = (cfg.abstract_en || '').trim().split('\n');
  const kw = (cfg.keywords_en || []).join(', ');
  return [
    pageBreak(),
    para('ABSTRACT', {
      size: SIZE_H1, bold: true,
      align: AlignmentType.CENTER,
      font: FONT_EN,
      spaceAfter: cm(0.3),
    }),
    mixedPara([
      { text: 'NAME：', bold: true, font: FONT_EN },
      { text: cfg.author_en, font: FONT_EN },
    ], { spaceBefore: 0, spaceAfter: 0 }),
    mixedPara([
      { text: 'MONTH/YEAR：', bold: true, font: FONT_EN },
      { text: `${cfg.month_en}, ${cfg.year_en}`, font: FONT_EN },
    ]),
    mixedPara([
      { text: 'ADVISER：', bold: true, font: FONT_EN },
      { text: cfg.advisors.map(a => `${a.name_en}, ${a.title_en}`).join('；'), font: FONT_EN },
    ]),
    mixedPara([
      { text: 'TITLE：', bold: true, font: FONT_EN },
      { text: cfg.title_en, font: FONT_EN },
    ], { spaceAfter: cm(0.3) }),
    ...lines.map(line => para(line, {
      indent: SIZE_BODY * 2,
      lineSpacing: 480,
      font: FONT_EN,
    })),
    ...emptyLine(1),
    mixedPara([
      { text: 'Keywords: ', bold: true, font: FONT_EN },
      { text: kw, font: FONT_EN },
    ]),
  ];
}

/** 目錄（自動） */
function buildTOC() {
  return [
    pageBreak(),
    para('目錄', {
      size: SIZE_H1, bold: true,
      align: AlignmentType.CENTER, spaceAfter: cm(0.5),
    }),
    new TableOfContents('目錄', {
      hyperlink: true,
      headingStyleRange: '1-2',
    }),
  ];
}

/** 章節內文 */
function buildChapters() {
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

/** 參考文獻 */
function buildReferences() {
  return [
    pageBreak(),
    para('參考文獻', {
      size: SIZE_H1, bold: true,
      align: AlignmentType.CENTER, spaceAfter: cm(0.5),
    }),
    ...(cfg.references || []).map(ref =>
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: 480, lineRule: LineRuleType.AUTO },
        indent: { left: SIZE_BODY * 2, hanging: SIZE_BODY * 2 },
        children: [
          new TextRun({ text: ref, size: SIZE_BODY, font: { name: isZh(ref) ? FONT_ZH : FONT_EN } }),
        ],
      })
    ),
  ];
}

// ─────────────────────────────────────────────────────────
// 組合文件
// ─────────────────────────────────────────────────────────
async function build() {
  const allChildren = [
    ...buildCoverPage(),
    ...buildApprovalPage(),
    ...buildAcknowledgements(),
    ...buildAbstractZh(),
    ...buildAbstractEn(),
    ...buildTOC(),
    ...buildChapters(),
    ...buildReferences(),
  ];

  const doc = new Document({
    creator: cfg.author_zh,
    title: cfg.title_zh,
    description: cfg.title_en,
    styles: {
      default: {
        document: {
          run: { font: { name: FONT_ZH }, size: SIZE_BODY },
        },
      },
      paragraphStyles: [
        {
          id: 'Heading1',
          name: 'Heading 1',
          basedOn: 'Normal',
          next: 'Normal',
          run: { size: SIZE_H1, bold: true, font: { name: FONT_ZH } },
          paragraph: {
            alignment: AlignmentType.CENTER,
            spacing: { before: cm(0.5), after: cm(0.5) },
          },
        },
        {
          id: 'Heading2',
          name: 'Heading 2',
          basedOn: 'Normal',
          next: 'Normal',
          run: { size: SIZE_H2, bold: true, font: { name: FONT_ZH } },
          paragraph: {
            spacing: { before: cm(0.3), after: cm(0.2) },
          },
        },
      ],
    },
    sections: [
      {
        properties: {
          page: {
            margin: {
              top: cm(3),
              right: cm(3),
              bottom: cm(3),
              left: cm(3.5),
            },
            size: {
              width: cm(21),
              height: cm(29.7),
            },
          },
          pageNumberStart: 1,
          pageNumberFormatType: NumberFormat.LOWER_ROMAN,
        },
        footers: {
          default: new Footer({
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                  new TextRun({ children: [PageNumber.CURRENT], font: { name: FONT_EN }, size: SIZE_BODY }),
                ],
              }),
            ],
          }),
        },
        children: allChildren,
      },
    ],
  });

  const outputPath = path.join(__dirname, cfg.output_filename || '碩士論文.docx');
  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(outputPath, buffer);
  console.log(`\n✅ 論文文件已產生：${outputPath}`);
  console.log(`   大小：${(buffer.length / 1024).toFixed(1)} KB`);
}

build().catch(err => {
  console.error('產生失敗：', err.message);
  process.exit(1);
});
