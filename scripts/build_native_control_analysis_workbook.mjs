#!/usr/bin/env node

import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";


function parseArgs(argv) {
  const args = {
    input: "outputs/native_control_analysis/workbook_payload.json",
    output: "outputs/native_control_analysis/作文母语对照分析结果.xlsx",
    previewDir: "outputs/native_control_analysis/workbook_previews",
  };
  for (let index = 2; index < argv.length; index += 1) {
    const key = argv[index];
    const value = argv[index + 1];
    if (key === "--input") args.input = value;
    if (key === "--output") args.output = value;
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


function safeSheetName(name, used) {
  let base = name.replace(/[\\/?*\[\]:]/g, "_").slice(0, 31);
  let candidate = base;
  let suffix = 2;
  while (used.has(candidate)) {
    const ending = `_${suffix}`;
    candidate = `${base.slice(0, 31 - ending.length)}${ending}`;
    suffix += 1;
  }
  used.add(candidate);
  return candidate;
}


function styleHeader(range) {
  range.format = {
    fill: "#365B8C",
    font: { bold: true, color: "#FFFFFF", size: 10 },
    verticalAlignment: "center",
    horizontalAlignment: "center",
    wrapText: true,
    borders: { preset: "outside", style: "thin", color: "#274664" },
  };
  range.format.rowHeight = 32;
}


function formatFor(header) {
  if (/作文编码|网页文章ID|正文SHA256/.test(header)) return "@";
  if (/占比|比例|覆盖率|显著$/.test(header)) return "0.00%";
  if (/p值/.test(header)) return "0.000E+00";
  if (/样本量|篇数|次数|变量数|因子数|成功次数|序号|迭代/.test(header)) return "#,##0";
  if (/均值|标准差|中位|偏差|效应|Hedges|Tucker|载荷|KMO|方差|系数|标准误|t值|R平方|得分|CI|特征根/.test(header)) return "0.000";
  return "General";
}


function widthFor(header) {
  if (/来源URL|说明/.test(header)) return 48;
  if (/字段名|维度|作文题目|项$/.test(header)) return 30;
  if (/作文文件名|网页文章ID/.test(header)) return 22;
  if (/诊断|筛选原因/.test(header)) return 36;
  return Math.min(22, Math.max(11, String(header).length * 1.4 + 2));
}


function addOverview(workbook, metadata) {
  const sheet = workbook.worksheets.add("研究概览");
  sheet.showGridLines = false;
  sheet.getRange("A1:H1").merge();
  sheet.getRange("A1").values = [["HSK学习者作文与公开网络母语参照语料对照分析"]];
  sheet.getRange("A1:H1").format = {
    fill: "#253F5B",
    font: { bold: true, color: "#FFFFFF", size: 16 },
    verticalAlignment: "center",
  };
  sheet.getRange("A1:H1").format.rowHeight = 36;
  sheet.getRange("A3:B3").values = [["分析项目", "结果"]];
  styleHeader(sheet.getRange("A3:B3"));
  const validation = metadata.validation || {};
  const joint = metadata.joint_factor_sensitivity || {};
  const rows = [
    ["学习者作文", validation.learner_rows || 0],
    ["母语参照作文", validation.native_rows || 0],
    ["宽表字段", validation.columns || 0],
    ["投影指标", validation.selected_features || 0],
    ["固定维度", 5],
    ["篇幅匹配重抽样", metadata.length_resamples || 0],
    ["Bootstrap次数", metadata.bootstrap_iterations || 0],
    ["联合平行分析次数", metadata.parallel_iterations || 0],
    ["联合建议因子数", joint.suggested_factors ?? "未得到"],
    ["联合选定因子数", joint.selected_factors ?? "未通过诊断"],
  ];
  sheet.getRangeByIndexes(3, 0, rows.length, 2).values = rows;
  sheet.getRange("D3:H3").merge();
  sheet.getRange("D3").values = [["解释边界"]];
  styleHeader(sheet.getRange("D3:H3"));
  sheet.getRange("D4:H9").merge();
  sheet.getRange("D4").values = [[
    "母语参照文本来自作文网高中栏目，四组各45篇且最终均为高中范围。栏目归类不能独立认证作者身份，也不能排除编辑加工；NJ2、NY1、NY2还大量使用宽泛主题，网页作文篇幅整体偏长。因此所有结论均表述为学习者语料与公开网络高中作文题目域参照语料之间的统计差异，不作母语身份、纯体裁或纯语言能力的因果解释。",
  ]];
  sheet.getRange("D4:H9").format = {
    fill: "#EEF3F7",
    font: { color: "#253238", size: 10 },
    wrapText: true,
    verticalAlignment: "center",
    borders: { preset: "outside", style: "thin", color: "#B8C6D0" },
  };
  sheet.getRange("D11:H11").merge();
  sheet.getRange("D11").values = [["版权处理"]];
  styleHeader(sheet.getRange("D11:H11"));
  sheet.getRange("D12:H14").merge();
  sheet.getRange("D12").values = [["网页全文仅本地保存并排除Git；本工作簿只包含来源元数据、派生统计和分析结果。"]];
  sheet.getRange("D12:H14").format = {
    fill: "#F8F4E9",
    font: { color: "#4A4030", size: 10 },
    wrapText: true,
    verticalAlignment: "center",
    borders: { preset: "outside", style: "thin", color: "#CDBD93" },
  };
  sheet.getRange("A4:A13").format.font = { bold: true, color: "#253F5B" };
  sheet.getRange("A1:A14").format.columnWidth = 22;
  sheet.getRange("B1:B14").format.columnWidth = 20;
  sheet.getRange("C1:C14").format.columnWidth = 3;
  sheet.getRange("D1:H14").format.columnWidth = 17;
  sheet.freezePanes.freezeRows(3);
}


function addDataSheet(workbook, requestedName, data, tableIndex, used) {
  const name = safeSheetName(requestedName, used);
  const sheet = workbook.worksheets.add(name);
  sheet.showGridLines = false;
  const headers = data.columns || [];
  const rows = data.rows || [];
  if (!headers.length) return sheet;
  sheet.getRangeByIndexes(0, 0, 1, headers.length).values = [headers];
  if (rows.length) sheet.getRangeByIndexes(1, 0, rows.length, headers.length).values = rows;
  styleHeader(sheet.getRangeByIndexes(0, 0, 1, headers.length));
  const lastColumn = columnName(headers.length - 1);
  const lastRow = rows.length + 1;
  if (rows.length) {
    const table = sheet.tables.add(`A1:${lastColumn}${lastRow}`, true, `NativeAnalysis${tableIndex}`);
    table.style = "TableStyleMedium2";
    table.showFilterButton = true;
  }
  for (let index = 0; index < headers.length; index += 1) {
    const letter = columnName(index);
    const range = sheet.getRange(`${letter}1:${letter}${lastRow}`);
    range.format.columnWidth = widthFor(headers[index]);
    if (/来源URL|说明|作文题目|字段名|维度|诊断/.test(headers[index])) {
      range.format.wrapText = true;
      range.format.verticalAlignment = "top";
    }
    const numberFormat = formatFor(headers[index]);
    if (numberFormat !== "General" && rows.length) {
      sheet.getRange(`${letter}2:${letter}${lastRow}`).format.numberFormat = numberFormat;
    }
  }
  const effectIndex = headers.findIndex((header) => /Hedges_g/.test(header));
  if (effectIndex >= 0 && rows.length) {
    const letter = columnName(effectIndex);
    sheet.getRange(`${letter}2:${letter}${lastRow}`).conditionalFormats.add("colorScale", {
      thresholds: [{ type: "min", value: -2 }, { type: "num", value: 0 }, { type: "max", value: 2 }],
      colors: ["#4A78B4", "#FFFFFF", "#C4493D"],
    });
  }
  sheet.freezePanes.freezeRows(1);
  return sheet;
}


async function main() {
  const args = parseArgs(process.argv);
  const payload = JSON.parse(await fs.readFile(args.input, "utf8"));
  const workbook = Workbook.create();
  addOverview(workbook, payload.metadata || {});
  const used = new Set(["研究概览"]);
  let tableIndex = 1;
  for (const [name, data] of Object.entries(payload.sheets || {})) {
    addDataSheet(workbook, name, data, tableIndex, used);
    tableIndex += 1;
  }
  await fs.mkdir(path.dirname(args.output), { recursive: true });
  await fs.mkdir(args.previewDir, { recursive: true });
  const targets = [["研究概览", "A1:H14"]];
  for (const [name, data] of Object.entries(payload.sheets || {})) {
    if (!(data.columns || []).length) continue;
    targets.push([name, `A1:${columnName(Math.min(data.columns.length - 1, 13))}${Math.min((data.rows || []).length + 1, 15)}`]);
  }
  for (const [sheetName, range] of targets) {
    try {
      const preview = await workbook.render({ sheetName, range, scale: 1, format: "png" });
      await fs.writeFile(path.join(args.previewDir, `${sheetName}.png`), new Uint8Array(await preview.arrayBuffer()));
    } catch (error) {
      console.warn(`Preview skipped for ${sheetName}: ${error.message}`);
    }
  }
  const inspection = await workbook.inspect({
    kind: "match",
    searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
    options: { useRegex: true, maxResults: 100 },
    summary: "final formula error scan",
  });
  console.log(inspection.ndjson);
  const output = await SpreadsheetFile.exportXlsx(workbook);
  await output.save(args.output);
  console.log(`Workbook written: ${args.output}`);
}


await main();
