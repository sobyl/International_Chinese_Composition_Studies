#!/usr/bin/env node

import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";


function parseArgs(argv) {
  const args = {
    input: "outputs/mfmd_analysis/workbook_payload.json",
    output: "outputs/mfmd_analysis/作文多维分析结果.xlsx",
    previewDir: "outputs/mfmd_analysis/workbook_previews",
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


function styleTitle(range) {
  range.format = {
    fill: "#253F5B",
    font: { bold: true, color: "#FFFFFF", size: 16 },
    verticalAlignment: "center",
    horizontalAlignment: "left",
  };
  range.format.rowHeight = 34;
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
  range.format.rowHeight = 30;
}


function numberFormatFor(header) {
  if (/作文编码/.test(header)) return "@";
  if (/样本量|次数|篇数|变量数|因子数|自由度$|序号$|成功次数|删除.*数/.test(header)) return "#,##0";
  if (/p值|校正p值/.test(header)) return "0.000E+00";
  if (/占比|比例|覆盖率/.test(header)) return "0.00%";
  if (/是否|通过|异常|大于/.test(header)) return "General";
  if (/KMO|MSA|相关|载荷|共同度|独特性|方差|均值|标准差|中位数|系数|误|t值|F$|Omega|Hedges|rho|R平方|得分|CI|特征根|稳定性/.test(header)) return "0.000";
  return "General";
}


function widthFor(header, rows, columnIndex) {
  const sample = rows.slice(0, 80).map((row) => row[columnIndex]);
  let length = String(header).length * 1.5;
  for (const value of sample) {
    if (value === null || value === undefined) continue;
    length = Math.max(length, String(value).length);
  }
  if (/作文片段/.test(header)) return 58;
  if (/筛选原因|诊断说明|正向|负向/.test(header)) return 34;
  if (/字段名|维度名称|主要维度|篇名|图名|文件/.test(header)) return Math.min(36, Math.max(18, length + 2));
  if (/作文编码|作文文件名/.test(header)) return 22;
  return Math.min(20, Math.max(10, length + 2));
}


function addOverview(workbook, payload) {
  const sheet = workbook.worksheets.add("研究概览");
  sheet.showGridLines = false;
  sheet.getRange("A1:H1").merge();
  sheet.getRange("A1").values = [["HSK作文语言特征 MF/MD 多维分析"]];
  styleTitle(sheet.getRange("A1:H1"));
  sheet.getRange("A3:B3").values = [["分析指标", "结果"]];
  styleHeader(sheet.getRange("A3:B3"));
  const metadata = payload.metadata;
  const selected = metadata.selected_model;
  const rows = [
    ["作文样本", metadata.validation.rows],
    ["篇名代码", "J1 / J2 / Y1 / Y2，各155篇"],
    ["初始候选变量", metadata.screening.initial_candidates],
    ["MSA筛选后变量", metadata.screening.after_msa],
    ["最终模型变量", selected.feature_count],
    ["最终维度数", selected.factor_count],
    ["KMO", selected.kmo],
    ["累计解释方差", selected.cumulative_variance],
    ["Bootstrap成功次数", selected.bootstrap_successes],
    ["模型选择说明", selected.selection_note],
  ];
  sheet.getRangeByIndexes(3, 0, rows.length, 2).values = rows;
  sheet.getRange("D3:H3").merge();
  sheet.getRange("D3").values = [["最终语言维度"]];
  styleHeader(sheet.getRange("D3:H3"));
  const dimensions = selected.dimension_labels.map((label, index) => [`D${index + 1}`, label]);
  sheet.getRangeByIndexes(3, 3, dimensions.length, 2).values = dimensions;
  sheet.getRange("D10:H10").merge();
  sheet.getRange("D10").values = [["解释边界"]];
  styleHeader(sheet.getRange("D10:H10"));
  sheet.getRange("D11:H14").merge();
  sheet.getRange("D11").values = [[
    "当前语料没有汉语母语者对照组；J/Y同时包含题目与体裁差异；国籍分布与题目存在混杂。报告中的维度与组间差异均为探索性统计关系，不作因果或母语者比较解释。",
  ]];
  sheet.getRange("D11:H14").format = {
    fill: "#EEF3F7",
    font: { color: "#253238", size: 10 },
    wrapText: true,
    verticalAlignment: "center",
    borders: { preset: "outside", style: "thin", color: "#B8C6D0" },
  };
  sheet.getRange("A4:A13").format.font = { bold: true, color: "#253F5B" };
  sheet.getRange("B4:B13").format.numberFormat = "0.000";
  sheet.getRange("B4:B9").format.numberFormat = "#,##0";
  sheet.getRange("B10").format.numberFormat = "0.000";
  sheet.getRange("B11").format.numberFormat = "0.0%";
  sheet.getRange("A1:H14").format.verticalAlignment = "center";
  sheet.getRange("A1:A14").format.columnWidth = 22;
  sheet.getRange("B1:B14").format.columnWidth = 32;
  sheet.getRange("C1:C14").format.columnWidth = 3;
  sheet.getRange("D1:D14").format.columnWidth = 10;
  sheet.getRange("E1:H14").format.columnWidth = 18;
  sheet.freezePanes.freezeRows(3);
  return sheet;
}


function addDataSheet(workbook, requestedName, sheetData, tableIndex, usedNames) {
  const sheetName = safeSheetName(requestedName, usedNames);
  const sheet = workbook.worksheets.add(sheetName);
  sheet.showGridLines = false;
  const columns = sheetData.columns;
  const rows = sheetData.rows;
  if (columns.length === 0) return sheet;
  sheet.getRangeByIndexes(0, 0, 1, columns.length).values = [columns];
  if (rows.length > 0) {
    sheet.getRangeByIndexes(1, 0, rows.length, columns.length).values = rows;
  }
  styleHeader(sheet.getRangeByIndexes(0, 0, 1, columns.length));
  sheet.freezePanes.freezeRows(1);
  const lastColumn = columnName(columns.length - 1);
  const lastRow = rows.length + 1;
  if (rows.length > 0) {
    const table = sheet.tables.add(`A1:${lastColumn}${lastRow}`, true, `MFMDTable${tableIndex}`);
    table.style = "TableStyleMedium2";
    table.showFilterButton = true;
  }
  for (let index = 0; index < columns.length; index += 1) {
    const letter = columnName(index);
    const columnRange = sheet.getRange(`${letter}1:${letter}${lastRow}`);
    columnRange.format.columnWidth = widthFor(columns[index], rows, index);
    const numberFormat = numberFormatFor(columns[index]);
    if (numberFormat !== "General" && rows.length > 0) {
      sheet.getRange(`${letter}2:${letter}${lastRow}`).format.numberFormat = numberFormat;
    }
    if (/作文片段|筛选原因|诊断说明|文件/.test(columns[index])) {
      columnRange.format.wrapText = true;
      columnRange.format.verticalAlignment = "top";
    }
  }
  if (requestedName === "因子载荷") {
    const loadingColumns = columns
      .map((header, index) => ({ header, index }))
      .filter(({ header }) => /^D\d+_/.test(header));
    for (const { index } of loadingColumns) {
      const letter = columnName(index);
      sheet.getRange(`${letter}2:${letter}${lastRow}`).conditionalFormats.add("colorScale", {
        thresholds: [
          { type: "min", value: -1 },
          { type: "num", value: 0 },
          { type: "max", value: 1 },
        ],
        colors: ["#4A78B4", "#FFFFFF", "#C4493D"],
      });
    }
  }
  return sheet;
}


async function main() {
  const args = parseArgs(process.argv);
  const payload = JSON.parse(await fs.readFile(args.input, "utf8"));
  const workbook = Workbook.create();
  const usedNames = new Set(["研究概览"]);
  addOverview(workbook, payload);
  let tableIndex = 1;
  for (const [name, data] of Object.entries(payload.sheets)) {
    addDataSheet(workbook, name, data, tableIndex, usedNames);
    tableIndex += 1;
  }

  await fs.mkdir(path.dirname(args.output), { recursive: true });
  await fs.mkdir(args.previewDir, { recursive: true });
  const previewTargets = [["研究概览", "A1:H14"]];
  for (const [name, data] of Object.entries(payload.sheets)) {
    const lastColumn = columnName(Math.max(0, data.columns.length - 1));
    const lastRow = Math.min(data.rows.length + 1, 16);
    previewTargets.push([name, `A1:${lastColumn}${lastRow}`]);
  }
  for (const [sheetName, range] of previewTargets) {
    try {
      const preview = await workbook.render({ sheetName, range, scale: 1, format: "png" });
      const bytes = new Uint8Array(await preview.arrayBuffer());
      await fs.writeFile(path.join(args.previewDir, `${sheetName}.png`), bytes);
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
