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
├── linguistic_features.py
├── resources/
│   └── 语言特征词表.csv
├── tests/
│   └── test_linguistic_features.py
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
- `outputs/作文词性统计宽表.xlsx`：基于 `clean_text` 生成的语言特征宽表。`词性统计` sheet 每行对应一篇作文，`字段说明` sheet 逐列记录定义、公式和分母。
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
- `segment_hsk_clean_texts.py`：使用 PyNLPIR 对 `clean_text` 分词，生成 `seg_text` 和语言特征宽表；细粒度词性用于语法特征，大类中文词性继续写入分词文本。
- `linguistic_features.py`：不依赖 PyNLPIR 运行环境的纯计算模块，负责词汇、句段、语法、篇章、记叙描写和 HSK 派生特征。
- `resources/语言特征词表.csv`：可人工审查的语言特征词表，字段为 `大类、特征名、词项、允许词性前缀、来源说明`。
- `tests/test_linguistic_features.py`：纯函数单元测试，覆盖 MATTR、句段切分、词表最长匹配、连续动词和 HSK 派生统计等口径。

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

默认语言特征词表为 `resources/语言特征词表.csv`，可通过 `--feature-lexicon` 替换。MATTR 窗口默认 50，可用 `--mattr-window` 调整；长句默认指汉字数严格大于 30，可用 `--long-sentence-threshold` 调整。

文本统计包含 `字数`、`纯文本字数`、`分词数`、`非标点分词数`和`去重词数`。其中，`去重词数`按 PyNLPIR 分词后的词形精确去重，排除标点，但保留数字和字母等非标点 token。

同形词对应多个主等级时，脚本优先使用 PyNLPIR 词性匹配词表词性；无法唯一判断时归入最低主等级。未命中词表的 token 不计入等级，但仍保留在占比分母中。

宽表按以下类别排列：基本信息、基础篇幅、词汇丰富度、词汇密度与词长、句段结构、词性、语法标记、篇章连接、记叙描写、HSK。次数型语言特征同时给出每千汉字频率，统一公式为 `次数 ÷ 纯文本字数 × 1000`；比例、均值、TTR、Guiraud 和 MATTR 直接保存计算值。非 HSK 词仅按专名、数字、字母串和其他拆分，其中“其他”不等同于错误词或高级词。

运行纯函数测试：

```bash
python -m unittest discover -s tests -v
```

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
       + resources/语言特征词表.csv
       + linguistic_features.py
  -> seg_text/
  -> outputs/作文词性统计宽表.xlsx
```

`作文样本主表.xlsx` 是后续分析的主索引；`clean_text` 是后续文本分析的正文来源；`seg_text` 和词性统计宽表是分词后的派生结果。
