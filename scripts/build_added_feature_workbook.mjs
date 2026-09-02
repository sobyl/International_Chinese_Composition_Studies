#!/usr/bin/env node

import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const DEFAULT_INPUT = "作文词性统计宽表.xlsx";
const DEFAULT_CATALOG = "语言特征说明.xlsx";
const DEFAULT_OUTPUT = "作文新增语言特征统计宽表.xlsx";

const BASIC_FIELDS = [
  "篇名代码", "篇名", "作文编码", "国籍", "作文题目", "作文分数", "体裁", "作文文件名",
];

const ADDED_FEATURE_NAMES = [
  "篇内高频词前10位",
  "汉字词分词数", "单音节词", "双音节词", "三音节及以上词",
  "平均段落长度_词", "标点占分词比例",
  "人名数", "地名数", "机构团体数", "其他专名数", "性质形容词数", "处所词数", "前缀数", "简称略语数",
  "成语数", "歇后语数", "惯用语数", "谚语数", "熟语多样性",
  "疑问副词数", "限定词数", "低调词与模糊语数", "夸张与加强语数",
  "表可能性情态词数", "表必要性情态词数", "表意愿性情态词数", "表语性情态动词数",
  "其他代词数", "其他助词数", "形容词作状语数", "特殊疑问句数", "是非问句数", "感叹句数", "把字句数", "被字句数",
  "因果复句数", "转折复句数", "条件复句数", "假设复句数", "目的复句数", "递进复句数", "并列复句数", "承接复句数", "解说复句数",
  "复句句次总数", "含关系标记复句数", "复句类型数", "复句类型多样性",
  "过去时间名词数", "现在时间名词数", "未来时间名词数", "其他时间名词数",
  "表完成时态时间副词数", "表过去时间副词数", "表正在进行时间副词数", "表将来时态时间副词数", "其他副词数",
  "私人性动词数", "建议要求类动词数", "公共性动词数", "动词是数", "其他动词数",
];

const CATEGORY_COLORS = {
  基本信息: "#375623",
  词汇丰富度: "#8064A2",
  词汇集中度: "#6F5A8A",
  词汇密度与词长: "#C65911",
  句段结构: "#BF9000",
  词性: "#548235",
  熟语: "#7A3E00",
  语法标记: "#2F75B5",
  复句关系: "#A64D79",
  记叙描写: "#7F6000",
};

function parseArgs(argv) {
  let input = DEFAULT_INPUT;
  let catalog = DEFAULT_CATALOG;
  let output = DEFAULT_OUTPUT;
  let previewDir = "tmp/added_feature_workbook";
  let dryRun = false;
  for (let index = 2; index < argv.length; index += 1) {
    const arg = argv[index];
    if (arg === "--input") input = argv[++index];
    else if (arg === "--catalog") catalog = argv[++index];
    else if (arg === "--output") output = argv[++index];
    else if (arg === "--preview-dir") previewDir = argv[++index];
    else if (arg === "--dry-run") dryRun = true;
    else throw new Error(`未知参数：${arg}`);
  }
  return { input, catalog, output, previewDir, dryRun };
}

function columnName(index) {
  let value = index;
  let result = "";
  while (value > 0) {
    value -= 1;
    result = String.fromCharCode(65 + (value % 26)) + result;
    value = Math.floor(value / 26);
  }
  return result;
}

function records(values, headerRow = 0) {
  const headers = values[headerRow];
  return values
    .slice(headerRow + 1)
    .filter((row) => row.some((value) => value !== null && value !== ""))
    .map((row) => Object.fromEntries(headers.map((header, index) => [header, row[index] ?? ""])));
}

function catalogFeatureFields(values) {
  const headerRow = values.findIndex((row) => row.includes("特征中文名"));
  if (headerRow < 0) throw new Error("语言特征说明中未找到表头");
  const wanted = new Set(ADDED_FEATURE_NAMES);
  const found = new Set();
  const fields = new Set();
  for (const row of records(values, headerRow)) {
    if (!wanted.has(row["特征中文名"])) continue;
    found.add(row["特征中文名"]);
    for (const field of String(row["宽表对应字段"]).split("；")) {
      if (field) fields.add(field);
    }
  }
  const missing = ADDED_FEATURE_NAMES.filter((name) => !found.has(name));
  if (missing.length) throw new Error(`语言特征说明缺少项目：${missing.join("、")}`);
  return fields;
}

function normalizedCategory(fieldName, category) {
  return fieldName.startsWith("篇内高频词前10位") ? "词汇集中度" : category;
}

function improvedDefinition(fieldName, definition) {
  if (fieldName === "篇内高频词前10位次数") {
    return "篇内出现频率最高的10个非标点词形的token次数之和；依据胡显耀TOP10概念改造为篇级词汇集中度指标";
  }
  if (fieldName === "篇内高频词前10位每千字") {
    return "篇内最高频10个词形的出现次数按每千汉字标准化；属于篇级改造指标";
  }
  if (fieldName === "篇内高频词前10位占比") {
    return "篇内最高频10个词形覆盖全部非标点token的比例；数值越高表示词汇使用越集中";
  }
  return definition;
}

function numberFormat(fieldName) {
  if (BASIC_FIELDS.includes(fieldName)) return fieldName === "作文分数" ? "0" : "@";
  if (/占比|比例/.test(fieldName)) return "0.00%";
  if (/多样性/.test(fieldName)) return "0.0000";
  if (/每千字|平均/.test(fieldName)) return "0.000";
  return "0";
}

function styleHeaderRuns(sheet, headers, categories) {
  let start = 0;
  while (start < headers.length) {
    const category = categories[start];
    let end = start;
    while (end + 1 < headers.length && categories[end + 1] === category) end += 1;
    sheet.getRange(`${columnName(start + 1)}1:${columnName(end + 1)}1`).format = {
      fill: CATEGORY_COLORS[category] || "#5B6573",
      font: { bold: true, color: "#FFFFFF" },
      horizontalAlignment: "center",
      verticalAlignment: "center",
      wrapText: true,
      borders: { preset: "all", style: "thin", color: "#D9D9D9" },
    };
    start = end + 1;
  }
}

async function buildWorkbook({ sourceValues, fieldRows, selectedHeaders, output, previewDir }) {
  const sourceHeaders = sourceValues[0];
  const sourceIndex = new Map(sourceHeaders.map((header, index) => [header, index]));
  const selectedValues = sourceValues.map((row, rowIndex) => selectedHeaders.map((header) => {
    const value = row[sourceIndex.get(header)] ?? null;
    if (rowIndex > 0 && header === "作文编码" && value !== null) return String(value);
    return value;
  }));
  const fieldByName = new Map(fieldRows.map((row) => [row["字段名"], row]));
  const selectedFieldRows = selectedHeaders.map((header, index) => {
    const source = fieldByName.get(header);
    if (!source) throw new Error(`字段说明缺少字段：${header}`);
    return [
      index + 1,
      header,
      normalizedCategory(header, source["类别"]),
      improvedDefinition(header, source["定义"]),
      source["公式"],
      source["单位/分母"],
    ];
  });
  const categories = selectedFieldRows.map((row) => row[2]);

  const workbook = Workbook.create();
  const dataSheet = workbook.worksheets.add("新增特征统计");
  dataSheet.showGridLines = false;
  dataSheet.getRangeByIndexes(0, 0, selectedValues.length, selectedHeaders.length).values = selectedValues;
  const lastColumn = columnName(selectedHeaders.length);
  const lastRow = selectedValues.length;
  const table = dataSheet.tables.add(`A1:${lastColumn}${lastRow}`, true, "AddedFeatureStatisticsTable");
  table.style = "TableStyleMedium4";
  table.showBandedRows = true;
  table.showFilterButton = true;
  styleHeaderRuns(dataSheet, selectedHeaders, categories);
  dataSheet.getRange(`A1:${lastColumn}1`).format.rowHeight = 44;
  dataSheet.getRange(`A2:${lastColumn}${lastRow}`).format = {
    verticalAlignment: "center",
    borders: { preset: "inside", style: "thin", color: "#E7E6E6" },
  };
  [11, 28, 23, 17, 30, 11, 13, 19].forEach((width, index) => {
    dataSheet.getRangeByIndexes(0, index, lastRow, 1).format.columnWidth = width;
  });
  dataSheet.getRange(`I:${lastColumn}`).format.columnWidth = 17;
  for (let index = 0; index < selectedHeaders.length; index += 1) {
    const column = dataSheet.getRangeByIndexes(1, index, lastRow - 1, 1);
    column.format.numberFormat = numberFormat(selectedHeaders[index]);
    if (index >= BASIC_FIELDS.length || selectedHeaders[index] === "作文分数") {
      column.format.horizontalAlignment = "right";
    }
  }
  dataSheet.getRange(`C2:C${lastRow}`).format.numberFormat = "@";
  dataSheet.freezePanes.freezeRows(1);
  dataSheet.freezePanes.freezeColumns(BASIC_FIELDS.length);

  const fieldSheet = workbook.worksheets.add("字段说明");
  fieldSheet.showGridLines = false;
  const fieldHeaders = ["序号", "字段名", "类别", "定义", "公式", "单位/分母"];
  const fieldValues = [fieldHeaders, ...selectedFieldRows];
  fieldSheet.getRangeByIndexes(0, 0, fieldValues.length, fieldHeaders.length).values = fieldValues;
  const fieldTable = fieldSheet.tables.add(`A1:F${fieldValues.length}`, true, "AddedFeatureDictionaryTable");
  fieldTable.style = "TableStyleMedium4";
  fieldTable.showBandedRows = true;
  fieldTable.showFilterButton = true;
  fieldSheet.getRange("A1:F1").format = {
    fill: "#385723",
    font: { bold: true, color: "#FFFFFF" },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
  };
  fieldSheet.getRange("A1:F1").format.rowHeight = 32;
  fieldSheet.getRange(`A2:F${fieldValues.length}`).format = {
    verticalAlignment: "top",
    wrapText: true,
    borders: { preset: "inside", style: "thin", color: "#E7E6E6" },
  };
  fieldSheet.getRange(`A2:F${fieldValues.length}`).format.rowHeight = 38;
  fieldSheet.getRange(`A2:A${fieldValues.length}`).format.horizontalAlignment = "center";
  for (let index = 0; index < selectedFieldRows.length; index += 1) {
    const category = selectedFieldRows[index][2];
    fieldSheet.getRangeByIndexes(index + 1, 2, 1, 1).format = {
      fill: CATEGORY_COLORS[category] || "#5B6573",
      font: { bold: true, color: "#FFFFFF" },
      verticalAlignment: "center",
    };
  }
  [9, 34, 20, 72, 64, 26].forEach((width, index) => {
    fieldSheet.getRangeByIndexes(0, index, fieldValues.length, 1).format.columnWidth = width;
  });
  fieldSheet.freezePanes.freezeRows(1);
  fieldSheet.freezePanes.freezeColumns(2);

  console.log((await workbook.inspect({
    kind: "table",
    range: "新增特征统计!A1:T8",
    include: "values,formulas",
    tableMaxRows: 8,
    tableMaxCols: 20,
    maxChars: 10000,
  })).ndjson);
  console.log((await workbook.inspect({
    kind: "match",
    searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
    options: { useRegex: true, maxResults: 300 },
    summary: "final formula error scan: added feature workbook",
  })).ndjson);

  const absoluteOutput = path.resolve(output);
  const temporaryOutput = path.join(path.dirname(absoluteOutput), `.${path.basename(absoluteOutput)}.tmp.xlsx`);
  const blob = await SpreadsheetFile.exportXlsx(workbook);
  await blob.save(temporaryOutput);
  await fs.rename(temporaryOutput, absoluteOutput);

  await fs.mkdir(previewDir, { recursive: true });
  const renderedWorkbook = await SpreadsheetFile.importXlsx(await FileBlob.load(absoluteOutput));
  const previews = [
    ["stats_start", "新增特征统计", "A1:T18"],
    ["stats_end", "新增特征统计", `${columnName(Math.max(1, selectedHeaders.length - 15))}1:${lastColumn}18`],
    ["dictionary", "字段说明", "A1:F30"],
  ];
  for (const [label, sheetName, range] of previews) {
    const preview = await renderedWorkbook.render({ sheetName, range, scale: 1.3, format: "png" });
    await fs.writeFile(path.join(previewDir, `${label}.png`), new Uint8Array(await preview.arrayBuffer()));
  }
  console.log(`Created ${absoluteOutput}: ${lastRow - 1} essays, ${selectedHeaders.length} columns`);
}

async function main() {
  const args = parseArgs(process.argv);
  const sourceWorkbook = await SpreadsheetFile.importXlsx(await FileBlob.load(path.resolve(args.input)));
  const catalogWorkbook = await SpreadsheetFile.importXlsx(await FileBlob.load(path.resolve(args.catalog)));
  const sourceValues = sourceWorkbook.worksheets.getItem("词性统计").getUsedRange(true).values;
  const fieldValues = sourceWorkbook.worksheets.getItem("字段说明").getUsedRange(true).values;
  const catalogValues = catalogWorkbook.worksheets.getItem("语言特征说明").getUsedRange(true).values;
  const fieldRows = records(fieldValues);
  const addedFieldSet = catalogFeatureFields(catalogValues);
  const sourceHeaders = sourceValues[0];
  const sourceHeaderSet = new Set(sourceHeaders);
  const missingBasics = BASIC_FIELDS.filter((field) => !sourceHeaderSet.has(field));
  if (missingBasics.length) throw new Error(`主宽表缺少基本信息字段：${missingBasics.join("、")}`);
  const addedHeaders = sourceHeaders.filter((field) => addedFieldSet.has(field));
  if (ADDED_FEATURE_NAMES.length !== 63) throw new Error(`新增语言特征项目应为63项，实际为${ADDED_FEATURE_NAMES.length}项`);
  if (addedHeaders.length !== 127) throw new Error(`新增统计字段应为127列，实际为${addedHeaders.length}列`);
  const selectedHeaders = [...BASIC_FIELDS, ...addedHeaders];
  if (args.dryRun) {
    console.log(`校验通过：${sourceValues.length - 1}篇作文，8列基本信息，63项新增语言特征，127个新增统计字段，共${selectedHeaders.length}列。`);
    return;
  }
  await buildWorkbook({ sourceValues, fieldRows, selectedHeaders, output: args.output, previewDir: args.previewDir });
}

await main();
