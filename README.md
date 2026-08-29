# International Chinese Composition Studies

本项目用于整理、抽样、抓取、清洗和分词 HSK 作文语料。当前主样本包含 `J1`、`J2`、`Y1`、`Y2` 四类作文，每类 155 篇，共 620 篇。

## 文件结构

```text
.
├── 4.14作文前筛.xlsx
├── 作文样本主表.xlsx
├── sample_essays_by_score.py
├── fetch_hsk_ori_texts.py
├── clean_hsk_ori_texts.py
├── segment_hsk_clean_texts.py
├── ori_text/
│   ├── J1/
│   ├── J2/
│   ├── Y1/
│   └── Y2/
├── clean_text/
│   ├── J1/
│   ├── J2/
│   ├── Y1/
│   └── Y2/
├── seg_text/
│   ├── J1/
│   ├── J2/
│   ├── Y1/
│   └── Y2/
└── outputs/
```

### 核心数据文件

- `4.14作文前筛.xlsx`：原始筛选表，包含多个 sheet，是后续抽样的源数据。
- `作文样本主表.xlsx`：当前主要迭代用表。四个 sheet 分别为 `J1`、`J2`、`Y1`、`Y2`，每个 sheet 155 篇。表内包含作文基本信息，并增加了 `作文文件名` 字段，例如 `J1_55_01`、`J1_60_02`。
- `outputs/新版HSK词汇大纲.csv`：从新版 HSK 考试大纲提取的 11,000 条机器可读词汇，包含主等级、兼属等级、词语、拼音和词性等字段。
- `outputs/作文词性统计宽表.xlsx`：基于 `clean_text` 分词后生成的统计宽表，每行对应一篇作文，包含词性数量以及 HSK 等级词汇次数和占比。
- `outputs/hsk_标注说明.md`：从 HSK 网站帮助页整理出的标注说明，用于理解和清洗原始标注文本。

### 文本目录

- `ori_text/{篇名代码}/{作文编码}.txt`：从 HSK 接口抓取的原始标注文本，保留网站返回的标注内容。
- `clean_text/{篇名代码}/{作文文件名}.txt`：清洗后的正文文本，只保留可确定的正确文本，保留原文目录不变。
- `seg_text/{篇名代码}/{作文文件名}.txt`：PyNLPIR 分词和词性标注结果。每行格式为 `词/中文词性 词/中文词性 ...`，保留原段落换行和标点。

当前三个文本目录均为 620 个 `.txt` 文件，其中 `J1/J2/Y1/Y2` 各 155 篇。

### 脚本

- `sample_essays_by_score.py`：按分数段抽样作文，尽量保持国籍多样和数量均衡，并排除中国香港、中国台湾、中国澳门和中国少数民族样本。
- `fetch_hsk_ori_texts.py`：批量从 HSK 网站接口抓取原始标注文本，保存到 `ori_text`。
- `clean_hsk_ori_texts.py`：清洗 `ori_text`，生成 `clean_text`，不修改原始文本。
- `segment_hsk_clean_texts.py`：使用 PyNLPIR 对 `clean_text` 分词，生成 `seg_text` 和 `outputs/作文词性统计宽表.xlsx`，并按 HSK 词表统计等级词汇。

## 常用运行命令

抓取原始文本：

```bash
python fetch_hsk_ori_texts.py
```

清洗原始文本：

```bash
python clean_hsk_ori_texts.py
```

分词并生成词性统计宽表：

```bash
arch -x86_64 /usr/bin/python3 segment_hsk_clean_texts.py
```

注意：当前 PyNLPIR 动态库只能在 x86_64 Python 下正常运行，所以分词脚本需要使用上面的 `arch -x86_64` 命令。

默认 HSK 词表为 `outputs/新版HSK词汇大纲.csv`，也可通过 `--hsk-vocab` 指定其他同结构 CSV。等级统计使用词表的 `主等级`：1-3 级合并为初等，4-6 级合并为中等，7-9 级为高等。每个等级的“词汇次数”按分词 token 出现次数统计，“词汇占比”的分母为该作文全部非标点分词数。

文本统计包含 `字数`、`纯文本字数`、`分词数`、`非标点分词数`和`去重词数`。其中，`去重词数`按 PyNLPIR 分词后的词形精确去重，排除标点，但保留数字和字母等非标点 token。

同形词对应多个主等级时，脚本优先使用 PyNLPIR 词性匹配词表词性；无法唯一判断时归入最低主等级。未命中词表的 token 不计入等级，但仍保留在占比分母中。

## 当前数据流

```text
4.14作文前筛.xlsx
  -> sample_essays_by_score.py
  -> 作文样本主表.xlsx
  -> fetch_hsk_ori_texts.py
  -> ori_text/
  -> clean_hsk_ori_texts.py
  -> clean_text/
  -> segment_hsk_clean_texts.py
       + outputs/新版HSK词汇大纲.csv
  -> seg_text/
  -> outputs/作文词性统计宽表.xlsx
```

`作文样本主表.xlsx` 是后续分析的主索引；`clean_text` 是后续文本分析的正文来源；`seg_text` 和词性统计宽表是分词后的派生结果。
