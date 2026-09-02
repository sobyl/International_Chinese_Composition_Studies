#!/usr/bin/env node

import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const DEFAULT_INPUTS = ["作文词性统计宽表.xlsx", "母语作文词性统计宽表.xlsx"];
const DEFAULT_OUTPUT = "语言特征说明.xlsx";
const SHEET_NAME = "语言特征说明";

const CATEGORY_COLORS = {
  基础篇幅: "#375623",
  词汇丰富度: "#8064A2",
  词汇集中度: "#6F5A8A",
  词汇密度与词长: "#C65911",
  句段结构: "#BF9000",
  词性: "#548235",
  熟语: "#7A3E00",
  语法标记: "#2F75B5",
  复句关系: "#A64D79",
  记叙描写: "#7F6000",
  HSK: "#5B9BD5",
};

const STATIC_EXAMPLES = {
  字数: "汉字、标点、数字和字母均计入",
  纯文本字数: "只计汉字：我爱汉语=4",
  分词数: "我/代词 爱/动词 汉语/名词 。/标点",
  非标点分词数: "上例排除句号后为3",
  去重词数: "学习、学习、汉语：2个类符",
  词形丰富度TTR: "100个形符、60个类符：TTR=0.60",
  Guiraud值: "类符数 V ÷ √形符数 N",
  "MATTR-50": "连续50个token构成一个滑动窗口",
  仅出现一次词: "全文只出现一次的词",
  篇内高频词前10位: "篇内频次最高的10个词",
  名词词汇多样性: "学校、老师、问题",
  动词词汇多样性: "学习、帮助、改变",
  形容词词汇多样性: "美丽、认真、雪白",
  副词词汇多样性: "很、已经、仍然",
  内容词数: "名词、动词、形容词、副词",
  实词数: "名词、动词、形容词、数词、量词等",
  虚词数: "副词、介词、连词、助词等",
  平均词长: "父母=2个汉字；国际化=3个汉字",
  汉字词分词数: "排除纯数字、字母和标点token",
  单音节词: "我、人、去",
  双音节词: "父母、学习、影响",
  三音节及以上词: "自行车、国际化",
  人名数: "鲁迅、小明、李老师",
  地名数: "北京、中国、长江",
  机构团体数: "教育部、北京大学、联合国",
  其他专名数: "专名标签中未归入人名、地名或机构者",
  名词数: "人、学校、问题",
  代词数: "我、你、他们、这",
  动词数: "学习、帮助、成为",
  形容词数: "好、雪白、大型",
  性质形容词数: "好、聪明、优秀",
  区别词数: "男、女、大型",
  状态词数: "雪白、通红、绿油油",
  副词数: "很、已经、仍然",
  介词数: "在、对、从、把",
  连词数: "和、但是、因为",
  助词数: "的、地、得、了",
  数词数: "一、两、三百",
  量词数: "个、本、次",
  时间词数: "今天、去年、晚上",
  处所词数: "这里、学校、家里",
  方位词数: "上、下、里面",
  前缀数: "第、老、阿",
  后缀数: "们、者、性",
  语气词数: "吧、呢、吗、啊",
  拟声词数: "轰、哗啦、叮咚",
  叹词数: "啊、哎、哦",
  简称略语数: "北大、世贸、人大",
  标点数: "，。！？；：",
  成语数: "一丝不苟、一帆风顺",
  歇后语数: "老王卖瓜——自卖自夸",
  惯用语数: "开绿灯、碰钉子",
  谚语数: "家和万事兴、远亲不如近邻",
  熟语数: "成语、歇后语、惯用语、谚语",
  熟语多样性: "熟语词种数 ÷ 熟语出现次数",
  时态助词了数: "看了、写了",
  时态助词着数: "看着、放着",
  时态助词过数: "去过、见过",
  结构助词的数: "美丽的城市",
  结构助词地数: "认真地学习",
  结构助词得数: "说得很好",
  介词把数: "把书放在桌上",
  介词被数: "被老师表扬",
  介词对数: "对孩子负责",
  介词给数: "给朋友写信",
  介词从数: "从北京出发",
  介词向数: "向老师请教",
  语气词吧数: "走吧",
  语气词呢数: "你呢？",
  语气词吗数: "你好吗？",
  语气词啊数: "真漂亮啊！",
  形容词作状语数: "认真学习、高兴地说",
  特殊疑问句数: "你为什么来？",
  是非问句数: "你去吗？",
  感叹句数: "这里真美啊！",
  把字句数: "我把门关上了。",
  被字句数: "他被老师表扬了。",
  因果复句数: "因为下雨，所以没去。",
  转折复句数: "虽然很累，但是很高兴。",
  条件复句数: "只要努力，就会进步。",
  假设复句数: "如果有时间，我就去。",
  目的复句数: "为了健康，他开始戒烟。",
  递进复句数: "不但认真，而且耐心。",
  并列复句数: "一边听，一边记录。",
  承接复句数: "先准备，然后出发。",
  解说复句数: "原因如下：第一……第二……",
  直接引语数: "他说：“我明天回来。”",
  第一人称代词占人称代词比例: "我、我们等 ÷ 全部人称代词",
  连续动词结构数: "去商店买东西",
  最长连续动词序列: "想去看看：连续3个动词token",
  "1级词汇": "爱、爸爸、白天",
  "2级词汇": "爱好、白色、帮忙",
  "3级词汇": "阿姨、安静、安全",
  "4级词汇": "爱情、爱心、安排",
  "5级词汇": "爱护、安全带、岸边",
  "6级词汇": "案例、按摩、暗示",
  "7-9级词汇": "挨家挨户、哀求、癌症",
  初等词汇: "HSK 1—3级词汇",
  中等词汇: "HSK 4—6级词汇",
  高等词汇: "HSK 7—9级词汇",
  HSK词汇: "命中《新版HSK词汇大纲》的token",
  非HSK词汇: "未命中HSK词表的非标点token",
  非HSK专名数: "人名、地名、机构名等",
  非HSK数字数: "1995、20、3.5",
  非HSK字母串数: "HSK、TV、Internet",
  非HSK其他数: "未归入专名、数字或字母串者",
};

const DESCRIPTION_OVERRIDES = {
  篇内高频词前10位: "篇内最高频10个非标点词形的覆盖程度；依据胡显耀TOP10概念改造为篇级词汇集中度指标，非原论文的子语料库统计",
  名词词汇多样性: "名词token去重后的词种规模及其相对丰富度",
  动词词汇多样性: "动词token去重后的词种规模及其相对丰富度",
  形容词词汇多样性: "性质形容词、区别词和状态词合并后的词种规模及其相对丰富度",
  副词词汇多样性: "副词token去重后的词种规模及其相对丰富度",
  单音节词: "仅含一个汉字的非标点token；以一个汉字近似一个音节",
  双音节词: "仅含两个汉字的非标点token；以一个汉字近似一个音节",
  三音节及以上词: "含三个及以上汉字的非标点token；以一个汉字近似一个音节",
  形容词数: "性质形容词、区别词和状态词三类数量之和",
  熟语数: "成语、歇后语、惯用语和谚语的出现次数之和",
  熟语多样性: "篇内熟语词种规模及其与熟语出现次数的比值",
  复句句次总数: "九类复句关系的句次之和；同一句可同时计入不同关系类型",
  含关系标记复句数: "至少含一种复句关系标记的不同句子数量",
  复句类型多样性: "篇内实际出现的复句关系类型数除以九类关系总数",
  HSK词汇: "按《新版HSK词汇大纲》精确匹配的非标点token",
  非HSK词汇: "未命中《新版HSK词汇大纲》的非标点token；不直接解释为错误词或高级词",
};

function parseArgs(argv) {
  const inputs = [];
  let output = DEFAULT_OUTPUT;
  let previewDir = "tmp/language_feature_catalog/final";
  let dryRun = false;
  for (let index = 2; index < argv.length; index += 1) {
    const arg = argv[index];
    if (arg === "--input") inputs.push(argv[++index]);
    else if (arg === "--output") output = argv[++index];
    else if (arg === "--preview-dir") previewDir = argv[++index];
    else if (arg === "--dry-run") dryRun = true;
    else throw new Error(`未知参数：${arg}`);
  }
  return { inputs: inputs.length ? inputs : DEFAULT_INPUTS, output, previewDir, dryRun };
}

function parseCsv(text) {
  const rows = [];
  let row = [];
  let cell = "";
  let quoted = false;
  const source = text.replace(/^\uFEFF/, "");
  for (let index = 0; index < source.length; index += 1) {
    const char = source[index];
    if (quoted) {
      if (char === '"' && source[index + 1] === '"') {
        cell += '"';
        index += 1;
      } else if (char === '"') quoted = false;
      else cell += char;
    } else if (char === '"') quoted = true;
    else if (char === ",") {
      row.push(cell);
      cell = "";
    } else if (char === "\n") {
      row.push(cell.replace(/\r$/, ""));
      rows.push(row);
      row = [];
      cell = "";
    } else cell += char;
  }
  if (cell || row.length) {
    row.push(cell.replace(/\r$/, ""));
    rows.push(row);
  }
  return rows;
}

async function loadLexiconExamples() {
  const examples = new Map();
  for (const filename of ["resources/语言特征词表.csv", "resources/论文补充语言特征词表.csv"]) {
    const rows = parseCsv(await fs.readFile(filename, "utf8"));
    const [headers, ...data] = rows;
    const featureIndex = headers.indexOf("特征名");
    const termIndex = headers.indexOf("词项");
    for (const row of data) {
      const feature = row[featureIndex]?.trim();
      const term = row[termIndex]?.trim();
      if (!feature || !term) continue;
      if (!examples.has(feature)) examples.set(feature, []);
      const values = examples.get(feature);
      if (!values.includes(term) && values.length < 3) values.push(term);
    }
  }
  return examples;
}

function canonicalFeature(name, category, allNames) {
  if (category === "基本信息") return null;
  if (category === "词汇丰富度") {
    if (name.startsWith("仅出现一次词")) return "仅出现一次词";
    if (name.startsWith("篇内高频词前10位")) return "篇内高频词前10位";
    for (const part of ["名词", "动词", "形容词", "副词"]) {
      if (name.startsWith(part) && (name.includes("去重词") || name === `${part}TTR`)) return `${part}词汇多样性`;
    }
  }
  if (category === "词汇密度与词长") {
    for (const part of ["单音节词", "双音节词", "三音节及以上词"]) {
      if (name.startsWith(part)) return part;
    }
  }
  if (category === "句段结构" && name.startsWith("超过30字长句")) return "超过30字长句";
  if (category === "熟语" && (name.startsWith("熟语去重") || name === "熟语多样性")) return "熟语多样性";
  if (category === "HSK") {
    for (const level of ["1级", "2级", "3级", "4级", "5级", "6级", "7-9级", "初等", "中等", "高等"]) {
      if (name.startsWith(`${level}词汇`)) return `${level}词汇`;
    }
    if (name.startsWith("HSK词汇")) return "HSK词汇";
    if (name.startsWith("非HSK词汇")) return "非HSK词汇";
  }
  if (name.endsWith("每千字")) {
    const stem = name.slice(0, -3);
    const candidates = [stem, `${stem}数`];
    if (category === "HSK") candidates.push(`${stem}次数`, `${stem}种类数`);
    return candidates.find((candidate) => allNames.has(candidate)) || stem;
  }
  return name;
}

function lexiconFeatureName(name) {
  const stem = name.endsWith("数") ? name.slice(0, -1) : name;
  if (/^(因果|转折|条件|假设|目的|递进|并列|承接|解说)复句$/.test(stem)) return `${stem}标记`;
  return stem;
}

function exampleFor(name, lexiconExamples) {
  if (STATIC_EXAMPLES[name]) return STATIC_EXAMPLES[name];
  const terms = lexiconExamples.get(lexiconFeatureName(name));
  return terms?.length ? terms.join("、") : "—";
}

function reportForms(fields) {
  const names = fields.map((field) => field["字段名"]);
  const forms = [];
  if (names.some((name) => /(?:数|次数)$/.test(name))) forms.push("原始次数");
  if (names.some((name) => name.includes("每千字"))) forms.push("每千汉字频率");
  if (names.some((name) => /(占比|覆盖率|比例|密度|实词率)$/.test(name))) forms.push("比例");
  if (names.some((name) => /(种类数|去重词数|熟语去重数)$/.test(name))) forms.push("词种数");
  if (names.some((name) => /(TTR|多样性)$/.test(name))) forms.push("多样性指标");
  if (!forms.length) forms.push("直接指标");
  return [...new Set(forms)].join("、");
}

function buildCatalog(fieldRows, lexiconExamples) {
  const allNames = new Set(fieldRows.map((row) => row["字段名"]));
  const groups = new Map();
  for (const field of fieldRows) {
    const key = canonicalFeature(field["字段名"], field["类别"], allNames);
    if (!key) continue;
    const category = key === "篇内高频词前10位" ? "词汇集中度" : field["类别"];
    const groupKey = `${category}\u0000${key}`;
    if (!groups.has(groupKey)) groups.set(groupKey, { category, name: key, fields: [] });
    groups.get(groupKey).fields.push(field);
  }

  return [...groups.values()].map((group, index) => {
    const first = group.fields[0];
    const definition = (DESCRIPTION_OVERRIDES[group.name] || first["定义"] || "").replace(/[。；]+$/, "");
    return [
      group.category,
      index + 1,
      group.name,
      `${definition}。报告形式：${reportForms(group.fields)}。`,
      exampleFor(group.name, lexiconExamples),
      group.fields.map((field) => field["字段名"]).join("；"),
    ];
  });
}

function parseSheetNames(inspectResult) {
  return inspectResult.ndjson
    .split("\n")
    .filter(Boolean)
    .map((line) => JSON.parse(line).name)
    .filter(Boolean);
}

async function addCatalogSheet(workbook, catalogRows) {
  const sheetInfo = await workbook.inspect({ kind: "sheet", include: "id,name", maxChars: 6000 });
  if (parseSheetNames(sheetInfo).includes(SHEET_NAME)) workbook.worksheets.getItem(SHEET_NAME).delete();

  const sheet = workbook.worksheets.add(SHEET_NAME);
  sheet.showGridLines = false;
  sheet.mergeCells("A1:F1");
  sheet.getRange("A1").values = [["语言特征说明"]];
  sheet.mergeCells("A2:F2");
  sheet.getRange("A2").values = [[
    "本表按语言学项目汇总两份作文统计宽表的字段；“识别与统计口径”使用中文说明，同一项目的次数、每千汉字频率、占比、词种数或TTR合并为一行。精确公式及分母以宽表中的“字段说明”sheet为准。",
  ]];

  const headers = ["大类", "ID", "特征中文名", "识别与统计口径", "例（Example）", "宽表对应字段"];
  sheet.getRange("A4:F4").values = [headers];
  sheet.getRangeByIndexes(4, 0, catalogRows.length, headers.length).values = catalogRows;

  sheet.getRange("A1:F1").format = {
    fill: "#385723",
    font: { bold: true, color: "#FFFFFF", size: 16 },
    horizontalAlignment: "center",
    verticalAlignment: "center",
  };
  sheet.getRange("A1:F1").format.rowHeight = 30;
  sheet.getRange("A2:F2").format = {
    fill: "#E2F0D9",
    font: { color: "#404040", italic: true },
    wrapText: true,
    verticalAlignment: "center",
  };
  sheet.getRange("A2:F2").format.rowHeight = 34;
  sheet.getRange("A4:F4").format = {
    fill: "#548235",
    font: { bold: true, color: "#FFFFFF" },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
    borders: { preset: "all", style: "thin", color: "#D9EAD3" },
  };
  sheet.getRange("A4:F4").format.rowHeight = 32;

  const lastRow = catalogRows.length + 4;
  const dataRange = sheet.getRange(`A5:F${lastRow}`);
  dataRange.format = {
    verticalAlignment: "top",
    wrapText: true,
    borders: { preset: "all", style: "thin", color: "#D9D9D9" },
  };
  dataRange.format.rowHeight = 42;
  sheet.getRange(`B5:B${lastRow}`).format.horizontalAlignment = "center";
  sheet.getRange(`A5:A${lastRow}`).format.font = { bold: true, color: "#FFFFFF" };
  for (let index = 0; index < catalogRows.length; index += 1) {
    const category = catalogRows[index][0];
    sheet.getRangeByIndexes(index + 4, 0, 1, 1).format.fill = CATEGORY_COLORS[category] || "#666666";
  }

  const widths = [16, 8, 28, 72, 36, 76];
  widths.forEach((width, index) => {
    sheet.getRangeByIndexes(0, index, lastRow, 1).format.columnWidth = width;
  });
  sheet.freezePanes.freezeRows(4);
  sheet.freezePanes.freezeColumns(2);

  const table = sheet.tables.add(`A4:F${lastRow}`, true, "LanguageFeatureCatalogTable");
  table.style = "TableStyleMedium4";
  table.showBandedRows = true;
  table.showFilterButton = true;
  return { sheet, lastRow };
}

async function createCatalogWorkbook(outputPath, catalogRows, previewDir) {
  const absolutePath = path.resolve(outputPath);
  const workbook = Workbook.create();
  const { lastRow } = await addCatalogSheet(workbook, catalogRows);

  const check = await workbook.inspect({
    kind: "table",
    range: `${SHEET_NAME}!A1:F18`,
    include: "values,formulas",
    tableMaxRows: 18,
    tableMaxCols: 6,
    maxChars: 10000,
  });
  console.log(check.ndjson);
  const errors = await workbook.inspect({
    kind: "match",
    searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
    options: { useRegex: true, maxResults: 300 },
    summary: `final formula error scan: ${path.basename(outputPath)}`,
  });
  console.log(errors.ndjson);

  await fs.mkdir(previewDir, { recursive: true });
  for (const [label, range] of [["top", "A1:F30"], ["middle", "A70:F100"], ["bottom", `A${Math.max(5, lastRow - 24)}:F${lastRow}`]]) {
    const preview = await workbook.render({ sheetName: SHEET_NAME, range, scale: 1.3, format: "png" });
    const outputName = `${path.parse(outputPath).name}_${label}.png`;
    await fs.writeFile(path.join(previewDir, outputName), new Uint8Array(await preview.arrayBuffer()));
  }

  const temporaryPath = path.join(path.dirname(absolutePath), `.${path.basename(absolutePath)}.tmp.xlsx`);
  const output = await SpreadsheetFile.exportXlsx(workbook);
  await output.save(temporaryPath);
  await fs.rename(temporaryPath, absolutePath);
  console.log(`Created ${absolutePath}: ${catalogRows.length} language features`);
}

async function removeCatalogSheet(inputPath) {
  const absolutePath = path.resolve(inputPath);
  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(absolutePath));
  const sheetInfo = await workbook.inspect({ kind: "sheet", include: "id,name", maxChars: 6000 });
  const sheetNames = parseSheetNames(sheetInfo);
  if (!sheetNames.includes(SHEET_NAME)) {
    console.log(`Skipped ${absolutePath}: no ${SHEET_NAME} sheet`);
    return;
  }
  workbook.worksheets.getItem(SHEET_NAME).delete();

  const remainingSheets = parseSheetNames(
    await workbook.inspect({ kind: "sheet", include: "id,name", maxChars: 6000 }),
  );
  if (remainingSheets.includes(SHEET_NAME)) throw new Error(`${inputPath} 未能移除 ${SHEET_NAME} sheet`);

  const errors = await workbook.inspect({
    kind: "match",
    searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
    options: { useRegex: true, maxResults: 300 },
    summary: `final formula error scan: ${path.basename(inputPath)}`,
  });
  console.log(errors.ndjson);

  const temporaryPath = path.join(path.dirname(absolutePath), `.${path.basename(absolutePath)}.catalog.tmp.xlsx`);
  const output = await SpreadsheetFile.exportXlsx(workbook);
  await output.save(temporaryPath);
  await fs.rename(temporaryPath, absolutePath);
  console.log(`Updated ${absolutePath}: removed ${SHEET_NAME} sheet`);
}

async function main() {
  const args = parseArgs(process.argv);
  const lexiconExamples = await loadLexiconExamples();
  const sourceWorkbook = await SpreadsheetFile.importXlsx(await FileBlob.load(path.resolve(args.inputs[0])));
  const values = sourceWorkbook.worksheets.getItem("字段说明").getUsedRange(true).values;
  const [headers, ...rows] = values;
  const fieldRows = rows
    .filter((row) => row.some((value) => value !== null && value !== ""))
    .map((row) => Object.fromEntries(headers.map((header, index) => [header, row[index] ?? ""])));
  const catalogRows = buildCatalog(fieldRows, lexiconExamples);
  if (catalogRows.length !== 178) throw new Error(`语言特征目录应为178项，实际为${catalogRows.length}项`);
  if (args.dryRun) {
    console.log(`校验通过：${catalogRows.length}项语言特征，将生成 ${args.output} 并清理${args.inputs.length}个宽表。`);
    return;
  }
  await createCatalogWorkbook(args.output, catalogRows, args.previewDir);
  for (const input of args.inputs) await removeCatalogSheet(input);
}

await main();
