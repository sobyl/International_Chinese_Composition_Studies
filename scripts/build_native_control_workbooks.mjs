#!/usr/bin/env node

import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";


const CATEGORY_COLORS = {
  "基本信息": "#365B8C",
  "基础篇幅": "#4B718C",
  "词汇丰富度": "#2F7D6D",
  "词汇密度与词长": "#527A3B",
  "句段结构": "#A56A27",
  "词性": "#7E5B95",
  "语法标记": "#A04755",
  "篇章连接": "#B17A32",
  "记叙描写": "#3E7D8B",
  "HSK": "#B4473A",
};


function parseArgs(argv) {
  const args = {
    selectedJson: "outputs/native_control/selected_samples.json",
    statsPayload: "outputs/native_control/native_stats_payload.json",
    masterOutput: "母语作文样本主表.xlsx",
    statsOutput: "母语作文词性统计宽表.xlsx",
    previewDir: "outputs/native_control/workbook_previews",
  };
  for (let index = 2; index < argv.length; index += 1) {
    const key = argv[index];
    const value = argv[index + 1];
    if (key === "--selected-json") args.selectedJson = value;
    if (key === "--stats-payload") args.statsPayload = value;
    if (key === "--master-output") args.masterOutput = value;
    if (key === "--stats-output") args.statsOutput = value;
    if (key === "--preview-dir") args.previewDir = value;
    if (key.startsWith("--")) index += 1;
  }
  return args;
}


function columnName(index) {
  let value = index + 1;
  let result = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    value = Math.floor((value - 1) / 26);
  }
  return result;
}


function styleHeader(range, fill = "#365B8C") {
  range.format = {
    fill,
    font: { bold: true, color: "#FFFFFF", size: 10 },
    verticalAlignment: "center",
    horizontalAlignment: "center",
    wrapText: true,
    borders: { preset: "outside", style: "thin", color: "#274664" },
  };
  range.format.rowHeight = 34;
}


function numberFormatFor(header) {
  if (/作文编码|网页文章ID|正文SHA256/.test(header)) return "@";
  if (/占比|比例|密度|覆盖率|TTR|目标篇幅相对偏差/.test(header)) return "0.00%";
  if (/p值|校正p值/.test(header)) return "0.000E+00";
  if (/日期/.test(header)) return "yyyy-mm-dd";
  if (/次数|词数|字数|分词数|句子数|段落数|标点数|序号|目标汉字数|正文汉字数/.test(header)) return "#,##0";
  if (/每千字|平均|中位|标准差|Guiraud|MATTR|偏差|最长|实虚词比/.test(header)) return "0.000";
  return "General";
}


function widthFor(header) {
  if (/来源URL/.test(header)) return 46;
  if (/作文题目|对应学习者篇名|来源类型|审核状态/.test(header)) return 28;
  if (/正文SHA256/.test(header)) return 26;
  if (/作文编码|网页文章ID|作文文件名/.test(header)) return 22;
  if (/定义|公式|单位/.test(header)) return 36;
  if (/字段名/.test(header)) return 28;
  return Math.min(20, Math.max(11, String(header).length * 1.5 + 2));
}


function addTable(sheet, headers, rows, tableName, categoryMap = null) {
  if (!headers.length) return;
  sheet.showGridLines = false;
  sheet.getRangeByIndexes(0, 0, 1, headers.length).values = [headers];
  if (rows.length) sheet.getRangeByIndexes(1, 0, rows.length, headers.length).values = rows;
  styleHeader(sheet.getRangeByIndexes(0, 0, 1, headers.length));
  if (categoryMap) {
    for (let index = 0; index < headers.length; index += 1) {
      const category = categoryMap.get(headers[index]) || "基本信息";
      const cell = sheet.getRangeByIndexes(0, index, 1, 1);
      cell.format.fill = CATEGORY_COLORS[category] || "#365B8C";
    }
  }
  const lastColumn = columnName(headers.length - 1);
  const lastRow = rows.length + 1;
  if (rows.length) {
    const table = sheet.tables.add(`A1:${lastColumn}${lastRow}`, true, tableName);
    table.style = "TableStyleMedium2";
    table.showFilterButton = true;
  }
  for (let index = 0; index < headers.length; index += 1) {
    const letter = columnName(index);
    const range = sheet.getRange(`${letter}1:${letter}${lastRow}`);
    range.format.columnWidth = widthFor(headers[index]);
    if (/来源URL|作文题目|定义|公式|单位|审核状态/.test(headers[index])) {
      range.format.wrapText = true;
      range.format.verticalAlignment = "top";
    }
    const format = numberFormatFor(headers[index]);
    if (format !== "General" && rows.length) {
      sheet.getRange(`${letter}2:${letter}${lastRow}`).format.numberFormat = format;
    }
  }
  sheet.freezePanes.freezeRows(1);
}


async function renderPreview(workbook, sheetName, range, outputPath) {
  try {
    const preview = await workbook.render({ sheetName, range, scale: 1, format: "png" });
    const bytes = new Uint8Array(await preview.arrayBuffer());
    await fs.writeFile(outputPath, bytes);
  } catch (error) {
    console.warn(`Preview skipped for ${sheetName}: ${error.message}`);
  }
}


async function scanAndExport(workbook, outputPath) {
  const inspection = await workbook.inspect({
    kind: "match",
    searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
    options: { useRegex: true, maxResults: 100 },
    summary: "final formula error scan",
  });
  console.log(inspection.ndjson);
  await fs.mkdir(path.dirname(path.resolve(outputPath)), { recursive: true });
  const output = await SpreadsheetFile.exportXlsx(workbook);
  await output.save(outputPath);
  console.log(`Workbook written: ${outputPath}`);
}


async function buildMaster(selectedDocument, outputPath, previewDir) {
  const rows = selectedDocument.selected || [];
  if (!rows.length) throw new Error("selected_samples.json 没有入选记录");
  const headers = Object.keys(rows[0]);
  const values = rows.map((row) => headers.map((header) => row[header] ?? ""));
  const workbook = Workbook.create();
  const sheet = workbook.worksheets.add("母语样本主表");
  addTable(sheet, headers, values, "NativeSampleTable");

  const note = workbook.worksheets.add("采样说明");
  note.showGridLines = false;
  note.getRange("A1:B1").merge();
  note.getRange("A1").values = [["公开网络母语参照语料采样说明"]];
  note.getRange("A1:B1").format = {
    fill: "#253F5B",
    font: { bold: true, color: "#FFFFFF", size: 16 },
    verticalAlignment: "center",
  };
  note.getRange("A1:B1").format.rowHeight = 36;
  const allHighSchool = rows.every((row) => row["学段"] === "高中");
  const notes = [
    ["语料定位", selectedDocument.corpus_label || "公开网络母语参照语料"],
    ["来源栏目", selectedDocument.source_site || "https://www.zuowen.com/gaozhong/"],
    ["服务协议", selectedDocument.source_agreement || "http://www.zuowen.com/help/agreement/"],
    ["版权处理", "网页全文仅保存在本地且不纳入Git；本表仅保留来源、审核和派生元数据。"],
    ["身份边界", "网站分类能证明文章被归入高中作文，不能独立核验作者身份或网站编辑程度。"],
    ["样本规模", `${rows.length}篇；${Object.entries(selectedDocument.summary || {}).map(([code, value]) => `${code}=${value.selected}`).join("，")}`],
    ["最终学段", allHighSchool ? "全部为高中范围；NY2初中最后备用通道未启用。" : "包含学段扩展样本，详见主表。"],
    ["来源口径", "完整的高中单元作文或练习来源作文可以纳入；仅含题目、要求、讲解或素材的页面排除。"],
    ["主题口径", "主题分为精确、近似、扩展、宽泛四级；宽泛样本属于相邻题目域参照，不等同于严格同题对照。"],
    ["篇幅限制", "高中网页作文整体长于学习者作文；目标篇幅偏差和后续稳健性检验结果须结合报告解读。"],
  ];
  const notesEndRow = notes.length + 3;
  note.getRange("A3:B3").values = [["项目", "说明"]];
  styleHeader(note.getRange("A3:B3"));
  note.getRangeByIndexes(3, 0, notes.length, 2).values = notes;
  note.getRange(`A4:A${notesEndRow}`).format.font = { bold: true, color: "#253F5B" };
  note.getRange(`A1:A${notesEndRow}`).format.columnWidth = 18;
  note.getRange(`B1:B${notesEndRow}`).format.columnWidth = 72;
  note.getRange(`B4:B${notesEndRow}`).format.wrapText = true;
  note.getRange(`A3:B${notesEndRow}`).format.verticalAlignment = "top";
  await renderPreview(workbook, "母语样本主表", `A1:${columnName(Math.min(headers.length - 1, 11))}${Math.min(rows.length + 1, 16)}`, path.join(previewDir, "母语样本主表.png"));
  await renderPreview(workbook, "采样说明", `A1:B${notesEndRow}`, path.join(previewDir, "采样说明.png"));
  await scanAndExport(workbook, outputPath);
}


async function buildStats(payload, outputPath, previewDir) {
  if ((payload.headers || []).length !== 301) throw new Error(`统计宽表应有301列，实际${(payload.headers || []).length}`);
  const workbook = Workbook.create();
  const categoryMap = new Map((payload.field_dictionary || []).map((row) => [row["字段名"], row["类别"]]));
  const stats = workbook.worksheets.add("词性统计");
  addTable(stats, payload.headers, payload.rows, "NativeStatsTable", categoryMap);

  const dictionaryRows = payload.field_dictionary || [];
  const dictionaryHeaders = dictionaryRows.length ? Object.keys(dictionaryRows[0]) : [];
  const dictionaryValues = dictionaryRows.map((row) => dictionaryHeaders.map((header) => row[header] ?? ""));
  const dictionary = workbook.worksheets.add("字段说明");
  addTable(dictionary, dictionaryHeaders, dictionaryValues, "NativeFieldDictionary");

  await renderPreview(workbook, "词性统计", "A1:V14", path.join(previewDir, "母语词性统计_前部.png"));
  await renderPreview(workbook, "词性统计", "HX1:IM14", path.join(previewDir, "母语词性统计_HSK区.png"));
  await renderPreview(workbook, "字段说明", "A1:F18", path.join(previewDir, "母语字段说明.png"));
  await scanAndExport(workbook, outputPath);
}


async function main() {
  const args = parseArgs(process.argv);
  const selectedDocument = JSON.parse(await fs.readFile(args.selectedJson, "utf8"));
  const statsPayload = JSON.parse(await fs.readFile(args.statsPayload, "utf8"));
  await fs.mkdir(args.previewDir, { recursive: true });
  await buildMaster(selectedDocument, args.masterOutput, args.previewDir);
  await buildStats(statsPayload, args.statsOutput, args.previewDir);
}


await main();
