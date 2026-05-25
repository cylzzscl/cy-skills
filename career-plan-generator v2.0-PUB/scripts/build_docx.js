#!/usr/bin/env node
/**
 * career-plan-generator · docx 构建脚本
 * --------------------------------------------------
 * 用法：
 *   node build_docx.js plan_data.json output.docx
 *
 * plan_data.json 的结构（见文末 EXAMPLE_DATA 演示）：
 *   {
 *     "studentName": "张三",         // 可选，用于封面 / 页眉 / 收尾
 *     "subtitle": "—— 计算机专业 · 互联网大厂方向 ——",
 *     "epigraph": "...",            // 一句话引言，可空
 *     "generatedAt": "2026-05",     // 拟定日期
 *     "planRange": "2026 — 2030",
 *     "portrait": "200 字以内的学生画像速写",
 *     "chapters": [
 *       {
 *         "title": "第一章 个人现状盘点",
 *         "blocks": [ {type:"p", text:"..."},
 *                     {type:"h2", text:"..."},
 *                     {type:"h3", text:"..."},
 *                     {type:"bullet", items:["...","..."]},
 *                     {type:"numbered", items:["...","..."]},
 *                     {type:"quote", text:"..."},
 *                     {type:"table", header:["列1","列2"], rows:[["..","..."]]},
 *                     {type:"pagebreak"}
 *                   ]
 *       },
 *       ...
 *     ],
 *     "closingLetter": "200 字给学生的一段话"   // 7.E
 *   }
 *
 * 依赖：docx (npm install -g docx)
 * --------------------------------------------------
 */

const fs = require("fs");
const path = require("path");

let docx;
try {
  docx = require("docx");
} catch (e) {
  console.error("[build_docx] 缺少依赖：docx。请先运行 `npm install -g docx`。");
  process.exit(2);
}

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, LevelFormat, TabStopType, TabStopPosition,
  HeadingLevel, BorderStyle, WidthType, ShadingType,
  VerticalAlign, PageNumber, PageBreak
} = docx;

// ---------- 配置 ----------
const FONT = "Microsoft YaHei";
const COLOR = {
  H1: "1F4E79",
  H2: "2E75B6",
  H3: "404040",
  text: "262626",
  mute: "8C8C8C",
  rule: "B4C7E7",
  tableHead: "D9E2F3",
  quoteBg: "F2F7FB",
  quoteBar: "2E75B6",
};

// ---------- 元素构造器 ----------
function P(text, { bold = false, size = 22, align, color = COLOR.text, italics = false, spacingAfter = 80 } = {}) {
  return new Paragraph({
    alignment: align,
    spacing: { before: 0, after: spacingAfter, line: 360 },
    children: [new TextRun({ text, bold, size, font: FONT, color, italics })],
  });
}

function H1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 200 },
    children: [new TextRun({ text, bold: true, size: 36, font: FONT, color: COLOR.H1 })],
  });
}
function H2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 280, after: 160 },
    children: [new TextRun({ text, bold: true, size: 28, font: FONT, color: COLOR.H2 })],
  });
}
function H3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 220, after: 120 },
    children: [new TextRun({ text, bold: true, size: 24, font: FONT, color: COLOR.H3 })],
  });
}

function bulletItem(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { before: 40, after: 40, line: 340 },
    children: [new TextRun({ text, size: 22, font: FONT, color: COLOR.text })],
  });
}
function numberedItem(text) {
  return new Paragraph({
    numbering: { reference: "numbers", level: 0 },
    spacing: { before: 40, after: 40, line: 340 },
    children: [new TextRun({ text, size: 22, font: FONT, color: COLOR.text })],
  });
}

function quoteBlock(text) {
  return new Paragraph({
    spacing: { before: 120, after: 120, line: 360 },
    indent: { left: 360, right: 360 },
    border: { left: { style: BorderStyle.SINGLE, size: 24, color: COLOR.quoteBar, space: 8 } },
    shading: { fill: COLOR.quoteBg, type: ShadingType.CLEAR },
    children: [new TextRun({ text, italics: true, size: 22, font: FONT, color: COLOR.H3 })],
  });
}

const cellBorder = { style: BorderStyle.SINGLE, size: 4, color: COLOR.rule };
const cellBorders = { top: cellBorder, bottom: cellBorder, left: cellBorder, right: cellBorder };

function cell(text, { bold = false, fill, width, align } = {}) {
  return new TableCell({
    borders: cellBorders,
    width: width ? { size: width, type: WidthType.DXA } : undefined,
    shading: fill ? { fill, type: ShadingType.CLEAR } : undefined,
    margins: { top: 100, bottom: 100, left: 140, right: 140 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({
      alignment: align,
      children: [new TextRun({ text: String(text ?? ""), bold, size: 20, font: FONT, color: COLOR.text })],
    })],
  });
}

function buildTable(header, rows) {
  // 默认 A4 1英寸边距下内容宽度约 9026 DXA，平均分配列宽
  const totalWidth = 9026;
  const colCount = header.length;
  const colWidth = Math.floor(totalWidth / colCount);
  const columnWidths = Array(colCount).fill(colWidth);

  const headerRow = new TableRow({
    tableHeader: true,
    children: header.map(h => cell(h, { bold: true, fill: COLOR.tableHead, width: colWidth, align: AlignmentType.CENTER })),
  });
  const dataRows = rows.map(row => new TableRow({
    children: row.map(c => cell(c, { width: colWidth })),
  }));

  return new Table({
    width: { size: totalWidth, type: WidthType.DXA },
    columnWidths,
    rows: [headerRow, ...dataRows],
  });
}

// ---------- block 渲染分发器 ----------
function renderBlock(block) {
  switch (block.type) {
    case "p":         return [P(block.text, block.opts || {})];
    case "h1":        return [H1(block.text)];
    case "h2":        return [H2(block.text)];
    case "h3":        return [H3(block.text)];
    case "bullet":    return (block.items || []).map(bulletItem);
    case "numbered":  return (block.items || []).map(numberedItem);
    case "quote":     return [quoteBlock(block.text)];
    case "table":     return [buildTable(block.header || [], block.rows || []), P(" ")];
    case "pagebreak": return [new Paragraph({ children: [new PageBreak()] })];
    case "blank":     return [P(" ")];
    default:
      return [P("[未知块类型：" + block.type + "]", { color: "C00000" })];
  }
}

// ---------- 主文档构造 ----------
function buildDoc(planData) {
  const studentName  = planData.studentName  || "同学";
  const subtitle     = planData.subtitle     || "—— 定制版职业规划 ——";
  const epigraph     = planData.epigraph     || "";
  const generatedAt  = planData.generatedAt  || "";
  const planRange    = planData.planRange    || "";
  const portrait     = planData.portrait     || "";
  const closingLetter = planData.closingLetter || "";

  const children = [];

  // 封面
  children.push(
    new Paragraph({ spacing: { before: 1800, after: 200 }, alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: "职业规划成长计划书", bold: true, size: 56, font: FONT, color: COLOR.H1 })] }),
    new Paragraph({ spacing: { before: 200, after: 200 }, alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: subtitle, size: 28, font: FONT, color: COLOR.H2 })] }),
  );
  if (epigraph) {
    children.push(
      new Paragraph({ spacing: { before: 1200, after: 100 }, alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: `「${epigraph}」`, italics: true, size: 24, font: FONT, color: "595959" })] })
    );
  }
  children.push(
    new Paragraph({ spacing: { before: 1400, after: 80 }, alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: `署名：${studentName}`, size: 24, font: FONT, color: COLOR.text })] }),
  );
  if (planRange) children.push(
    new Paragraph({ spacing: { before: 80, after: 80 }, alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: `规划周期：${planRange}`, size: 24, font: FONT, color: COLOR.text })] }),
  );
  if (generatedAt) children.push(
    new Paragraph({ spacing: { before: 80, after: 80 }, alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: `拟定日期：${generatedAt}`, size: 24, font: FONT, color: COLOR.text })] }),
  );
  // v1.4 新增：封面底部声明小字
  children.push(
    new Paragraph({ spacing: { before: 2400, after: 60 }, alignment: AlignmentType.CENTER,
      children: [new TextRun({
        text: "本报告由国科职规体系 + AI 生成，受限于模型能力，仅供参看。",
        size: 16, font: FONT, color: "8C8C8C", italics: true
      })] }),
    new Paragraph({ spacing: { before: 20, after: 80 }, alignment: AlignmentType.CENTER,
      children: [new TextRun({
        text: "如有需求，可向国科职规老师获取专业能力测评与职规服务。",
        size: 16, font: FONT, color: "8C8C8C", italics: true
      })] }),
  );
  children.push(new Paragraph({ children: [new PageBreak()] }));

  // 学生画像速写
  if (portrait) {
    children.push(
      H1("学生画像速写"),
      quoteBlock(portrait),
      new Paragraph({ children: [new PageBreak()] })
    );
  }

  // 各章
  for (const chapter of (planData.chapters || [])) {
    children.push(H1(chapter.title || ""));
    for (const block of (chapter.blocks || [])) {
      for (const node of renderBlock(block)) children.push(node);
    }
    children.push(new Paragraph({ children: [new PageBreak()] }));
  }

  // 收尾那段话
  if (closingLetter) {
    children.push(
      H1("写给你的一封信"),
      quoteBlock(closingLetter),
      P(" "),
      P("—— 本计划书 v1.0 完 ——", { align: AlignmentType.CENTER, color: COLOR.mute, italics: true })
    );
  }

  const doc = new Document({
    creator: "career-plan-generator",
    title: `${studentName} - 职业规划成长计划书`,
    styles: {
      default: { document: { run: { font: FONT, size: 22 } } },
      paragraphStyles: [
        { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 36, bold: true, font: FONT, color: COLOR.H1 },
          paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0 } },
        { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 28, bold: true, font: FONT, color: COLOR.H2 },
          paragraph: { spacing: { before: 280, after: 160 }, outlineLevel: 1 } },
        { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 24, bold: true, font: FONT, color: COLOR.H3 },
          paragraph: { spacing: { before: 220, after: 120 }, outlineLevel: 2 } },
      ],
    },
    numbering: {
      config: [
        { reference: "bullets",
          levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
        { reference: "numbers",
          levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      ],
    },
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 }, // A4
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
        },
      },
      headers: {
        default: new Header({ children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [new TextRun({ text: `${studentName}的职业规划成长计划书 · v1.4`, size: 18, font: FONT, color: COLOR.mute })],
        })] }),
      },
      footers: {
        default: new Footer({ children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({ text: "— 第 ", size: 18, font: FONT, color: COLOR.mute }),
            new TextRun({ children: [PageNumber.CURRENT], size: 18, font: FONT, color: COLOR.mute }),
            new TextRun({ text: " 页 —", size: 18, font: FONT, color: COLOR.mute }),
          ],
        })] }),
      },
      children,
    }],
  });

  return doc;
}

// ---------- CLI 入口 ----------
async function main() {
  const [, , dataPath, outPath] = process.argv;
  if (!dataPath || !outPath) {
    console.error("用法：node build_docx.js <plan_data.json> <output.docx>");
    process.exit(1);
  }
  const raw = fs.readFileSync(dataPath, "utf-8");
  const planData = JSON.parse(raw);
  const doc = buildDoc(planData);
  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(outPath, buffer);
  console.log(`[build_docx] 已生成：${outPath}（${buffer.length} 字节）`);
}

// 示例数据（仅用于参考结构，不会被自动执行）
const EXAMPLE_DATA = {
  studentName: "张三",
  subtitle: "—— 计算机专业 · 互联网大厂方向 ——",
  epigraph: "凡事预则立，不预则废。",
  generatedAt: "2026-05",
  planRange: "2026 — 2030",
  portrait: "（200 字以内的学生画像速写……）",
  chapters: [
    {
      title: "第一章 个人现状盘点",
      blocks: [
        { type: "h2", text: "1.1 学业表现" },
        { type: "p",  text: "（具体到学生填写的绩点 / 排名…）" },
        { type: "bullet", items: ["绩点 3.85 / 4.0", "专业排名前 15%"] },
        { type: "table",
          header: ["维度", "现状", "评估"],
          rows: [["GPA", "3.85", "稳"], ["排名", "前 15%", "强"]] },
      ],
    },
  ],
  closingLetter: "（200 字给学生的一段话……）",
};

if (require.main === module) {
  main().catch(err => { console.error(err); process.exit(1); });
}

module.exports = { buildDoc, EXAMPLE_DATA };
