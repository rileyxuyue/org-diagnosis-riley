/**
 * report_builder.js
 * 从 stdin 读取 JSON payload，生成完整 docx 报告
 * 用法：node report_builder.js <output_path>
 */

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  ImageRun, AlignmentType, BorderStyle, WidthType, ShadingType,
  VerticalAlign, HeadingLevel, PageBreak,
} = require("docx");
const fs = require("fs");

// ── 颜色常量 ─────────────────────────────────────────────────
const C_BLUE      = "2E75B6";
const C_BLUE_LIGHT = "D6E4F0";
const C_BLUE_MID  = "4472C4";
const C_GRAY      = "F5F5F5";
const C_GRAY2     = "E8E8E8";
const C_WHITE     = "FFFFFF";
const C_RED       = "C00000";
const C_GREEN     = "375623";
const C_BLACK     = "000000";

// ── 边框 ──────────────────────────────────────────────────────
const bdr = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: bdr, bottom: bdr, left: bdr, right: bdr };
const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

// ── 工具函数 ──────────────────────────────────────────────────
function gap(n = 1) {
  return Array(n).fill(null).map(() =>
    new Paragraph({ children: [new TextRun({ text: "" })] })
  );
}

function heading(text, level, color = C_BLUE) {
  const sizes = { 1: 32, 2: 26, 3: 22 };
  return new Paragraph({
    spacing: { before: 200, after: 120 },
    children: [new TextRun({
      text, bold: true,
      size: sizes[level] || 22,
      color, font: "Arial",
    })]
  });
}

function para(text, opts = {}) {
  return new Paragraph({
    spacing: { before: 60, after: 60 },
    children: [new TextRun({
      text,
      size:   opts.size   || 20,
      bold:   opts.bold   || false,
      color:  opts.color  || C_BLACK,
      font:   opts.font   || "Arial",
      italics: opts.italic || false,
    })]
  });
}

function hCell(text, w, bg = C_BLUE_LIGHT) {
  return new TableCell({
    width: { size: w, type: WidthType.DXA },
    borders,
    shading: { fill: bg, type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text, bold: true, size: 18, font: "Arial", color: C_BLUE })]
    })]
  });
}

function dCell(text, w, opts = {}) {
  const align = opts.center ? AlignmentType.CENTER : AlignmentType.LEFT;
  const color = opts.color || C_BLACK;
  const bg    = opts.bg    || C_WHITE;
  return new TableCell({
    width: { size: w, type: WidthType.DXA },
    borders,
    shading: { fill: bg, type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({
      alignment: align,
      children: [new TextRun({
        text: String(text ?? "-"),
        size: opts.size || 18,
        bold: opts.bold || false,
        font: "Arial",
        color,
      })]
    })]
  });
}

// ── 模块1：核心维度表 ─────────────────────────────────────────
function buildCoreTable(coreData) {
  const TW = 9026;
  const cols = [1600, 2000, 2126, 1500, 1800];

  function rowColor(metric) {
    return metric === "敬业度" ? "EBF3FB" : C_WHITE;
  }

  return new Table({
    width: { size: TW, type: WidthType.DXA },
    columnWidths: cols,
    rows: [
      new TableRow({
        tableHeader: true,
        children: [
          hCell("指标",       cols[0]),
          hCell("在BG的排名", cols[1]),
          hCell("排名变化",   cols[2]),
          hCell("分值",       cols[3]),
          hCell("较去年变化", cols[4]),
        ]
      }),
      ...coreData.map(r => new TableRow({
        children: [
          dCell(r.metric,      cols[0], { bold: true, bg: rowColor(r.metric) }),
          dCell(r.in_bg_rank,  cols[1], { center: true, bg: rowColor(r.metric) }),
          dCell(r.rank_change, cols[2], { center: true, bg: rowColor(r.metric) }),
          dCell(r.score,       cols[3], { center: true, bg: rowColor(r.metric) }),
          dCell(r.yoy_change,  cols[4], { center: true, bg: rowColor(r.metric) }),
        ]
      }))
    ]
  });
}

// ── 模块2：风险区间表 ─────────────────────────────────────────
function buildRiskTable(riskData) {
  const TW = 9026;
  const cols = [1600, 1326, 1400, 2300, 2400];

  return new Table({
    width: { size: TW, type: WidthType.DXA },
    columnWidths: cols,
    rows: [
      new TableRow({
        tableHeader: true,
        children: [
          hCell("风险区间",    cols[0]),
          hCell("题数",        cols[1]),
          hCell("占比",        cols[2]),
          hCell("Great Boss",  cols[3]),
          hCell("其他",        cols[4]),
        ]
      }),
      ...riskData.map((r, i) => {
        const bg = i === 0 ? "FFF2CC" : "FCE4D6";
        return new TableRow({
          children: [
            dCell(r.zone,         cols[0], { bold: true, bg }),
            dCell(r.count,        cols[1], { center: true, bg }),
            dCell(r.pct,          cols[2], { center: true, bg }),
            dCell(r.count_gb,     cols[3], { center: true, bg }),
            dCell(r.count_other,  cols[4], { center: true, bg }),
          ]
        });
      })
    ]
  });
}

// ── 模块4：排名靠后明细表 ────────────────────────────────────
function buildBottomTable(bottomData) {
  const TW = 9026;
  const cols = [1600, 4626, 1300, 1500];

  if (!bottomData.length) {
    return para("该部门无末10%题目。", { italic: true, color: "888888" });
  }

  return new Table({
    width: { size: TW, type: WidthType.DXA },
    columnWidths: cols,
    rows: [
      new TableRow({
        tableHeader: true,
        children: [
          hCell("标题",     cols[0]),
          hCell("题目",     cols[1]),
          hCell("BG内排名\n（倒数）", cols[2]),
          hCell("分值变化", cols[3]),
        ]
      }),
      ...bottomData.map(r => new TableRow({
        children: [
          dCell(r.short,       cols[0], { bold: true }),
          dCell(r.full,        cols[1], { size: 17 }),
          dCell(r.bg_rank_str, cols[2], { center: true, color: C_RED }),
          dCell(r.yoy_change,  cols[3], { center: true }),
        ]
      }))
    ]
  });
}

// ── 模块5：细分项大表 ────────────────────────────────────────
function buildSubdivTable(subdivData) {
  const TW = 9026;
  const cols = [1600, 3826, 1000, 900, 800, 900];

  const GROUP_COLORS = {
    "bg_higher":   "EBF3FB",
    "dept_higher": "E2EFDA",
    "dept_lag":    "FCE4D6",
    "dept_lead":   "EBF3FB",
  };

  const GROUP_HEADERS = {
    "bg_higher":   "BG排名较高的题（这些题在BG整体表现好，但本部门相对落后，值得重点关注）",
    "dept_higher": "部门排名较高的题（这些题本部门表现优于BG整体，是部门的相对优势项）",
    "dept_lag":    "部门增幅落后BG（这些题BG整体涨幅明显，但本部门改善较慢，需警惕差距拉大）",
    "dept_lead":   "部门增幅领先BG（这些题本部门改善速度明显快于BG整体，是近期进步亮点）",
  };

  const GROUP_HEADER_COLORS = {
    "bg_higher":   "BDD7EE",
    "dept_higher": "C6EFCE",
    "dept_lag":    "F4B8A0",
    "dept_lead":   "BDD7EE",
  };

  if (!subdivData.length) {
    return para("该部门无符合条件的细分项。", { italic: true, color: "888888" });
  }

  const rows = [
    // 列标题行
    new TableRow({
      tableHeader: true,
      children: [
        hCell("标题",     cols[0]),
        hCell("题目",     cols[1]),
        hCell("部门排名", cols[2]),
        hCell("BG排名",   cols[3]),
        hCell("排名差",   cols[4]),
        hCell("增幅差",   cols[5]),
      ]
    })
  ];

  let lastGroup = null;
  subdivData.forEach(r => {
    const bg        = C_WHITE;
    const headerBg  = GROUP_HEADER_COLORS[r.group] || C_GRAY2;
    const isNewGroup = r.group !== lastGroup;
    lastGroup = r.group;

    // 分组标题行（合并全部6列）
    if (isNewGroup) {
      const headerText = GROUP_HEADERS[r.group] || r.group_label;
      // 括号前为粗体，括号内为普通字
      const parenIdx = headerText.indexOf("（");
      const boldPart  = parenIdx > -1 ? headerText.slice(0, parenIdx) : headerText;
      const lightPart = parenIdx > -1 ? headerText.slice(parenIdx)    : "";

      rows.push(new TableRow({
        children: [
          new TableCell({
            columnSpan: 6,
            width: { size: TW, type: WidthType.DXA },
            borders,
            shading: { fill: headerBg, type: ShadingType.CLEAR },
            margins: { top: 100, bottom: 100, left: 160, right: 160 },
            children: [new Paragraph({
              children: [
                new TextRun({ text: boldPart,  bold: true, size: 19, font: "Arial", color: C_BLACK }),
                new TextRun({ text: lightPart, bold: false, size: 18, font: "Arial", color: "555555" }),
              ]
            })]
          })
        ]
      }));
    }

    const rankDiffStr = r.rank_diff != null
      ? (r.rank_diff >= 0 ? `+${r.rank_diff}` : String(r.rank_diff))
      : "-";
    const rankColor = r.group === "bg_higher" ? C_RED :
                      r.group === "dept_higher" ? C_GREEN : C_BLACK;

    rows.push(new TableRow({
      children: [
        dCell(r.short,           cols[0], { bold: true, bg }),
        dCell(r.full,            cols[1], { bg, size: 16 }),
        dCell(r.dept_rank,       cols[2], { center: true, bg }),
        dCell(r.bg_rank,         cols[3], { center: true, bg }),
        dCell(rankDiffStr,       cols[4], { center: true, bg, color: rankColor }),
        dCell(r.growth_diff_str, cols[5], { center: true, bg }),
      ]
    }));
  });

  return new Table({
    width: { size: TW, type: WidthType.DXA },
    columnWidths: cols,
    rows,
  });
}

// ── 主逻辑 ────────────────────────────────────────────────────
async function main() {
  const outPath = process.argv[2];
  if (!outPath) {
    console.error("用法：node report_builder.js <output_path>");
    process.exit(1);
  }

  let raw = "";
  process.stdin.setEncoding("utf8");
  for await (const chunk of process.stdin) raw += chunk;
  const p = JSON.parse(raw);

  // 图表图片
  const chartBuf = Buffer.from(p.chart_b64, "base64");

  // ── 组装文档内容 ──────────────────────────────────────────
  const children = [

    // ── 标题页 ───────────────────────────────────────────────
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 600, after: 200 },
      children: [new TextRun({
        text: p.dept_name,
        bold: true, size: 44, color: C_BLUE, font: "Arial",
      })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 120 },
      children: [new TextRun({
        text: `所属BG：${p.bg}`,
        size: 24, color: "555555", font: "Arial",
      })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 600 },
      children: [new TextRun({
        text: `生成日期：${p.today}`,
        size: 22, color: "888888", font: "Arial",
      })]
    }),
    new Paragraph({ children: [new PageBreak()] }),

    // ── 2.3 敬满 ────────────────────────────────────────────
    heading("2.3  敬满", 1),

    // ── 2.3.1 总分与定位 ────────────────────────────────────
    heading("2.3.1  总分与定位", 2),
    para("以下数据反映本部门在所属BG内的敬业度及满意度排名情况。", { italic: true, color: "555555" }),
    ...gap(1),

    // 核心维度表
    para("核心维度概览", { bold: true }),
    ...gap(1),
    buildCoreTable(p.core_data),
    ...gap(1),

    // 风险区间表
    para("风险区间分布", { bold: true }),
    ...gap(1),
    buildRiskTable(p.risk_data),
    ...gap(1),

    // 柱状图
    para("47道敬满子题 BG内击败率分布", { bold: true }),
    ...gap(1),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new ImageRun({
        type: "png",
        data: chartBuf,
        transformation: { width: 580, height: 370 },
        altText: { title: "击败率柱状图", description: "47道敬满子题BG内击败率", name: "chart" },
      })]
    }),
    ...gap(1),

    // 排名靠后明细表
    para("BG末10%题目明细", { bold: true }),
    ...gap(1),
    buildBottomTable(p.bottom_data),
    ...gap(2),

    // ── 2.3.2 值得关注的细分项 ───────────────────────────────
    new Paragraph({ children: [new PageBreak()] }),
    heading("2.3.2  值得关注的细分项", 2),
    para("以下题目在部门排名与BG排名、或增幅之间存在显著差异（差值 > 10）。", { italic: true, color: "555555" }),
    ...gap(1),
    buildSubdivTable(p.subdiv_data),
    ...gap(2),

    // ── 2.3.3 敬满开放题 ────────────────────────────────────
    new Paragraph({ children: [new PageBreak()] }),
    heading("2.3.3  敬满开放题", 2),
    para("（内容待补充）", { italic: true, color: "AAAAAA" }),
  ];

  const doc = new Document({
    styles: {
      default: {
        document: { run: { font: "Arial", size: 20 } }
      }
    },
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1200, right: 1200, bottom: 1200, left: 1200 }
        }
      },
      children,
    }]
  });

  const buf = await Packer.toBuffer(doc);
  fs.writeFileSync(outPath, buf);
  console.log(`OK: ${outPath}`);
}

main().catch(e => { console.error(e); process.exit(1); });
