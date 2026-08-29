# International Chinese Composition Studies

本项目用于整理、抽样、抓取、清洗和分词 HSK 作文语料。当前主样本包含 `J1`、`J2`、`Y1`、`Y2` 四类作文，每类 155 篇，共 620 篇。

## 文件结构

```text
.
├── 4.14作文前筛.xlsx
├── 作文样本主表.xlsx
├── 作文词性统计宽表.xlsx
├── 作文语言特征多维分析报告.docx
├── 作文语言特征多维分析报告.pdf
├── requirements-analysis.txt
├── scripts/
│   ├── __init__.py
│   ├── sample_essays_by_score.py
│   ├── fetch_hsk_ori_texts.py
│   ├── clean_hsk_ori_texts.py
│   ├── extract_hsk_vocabulary.py
│   ├── segment_hsk_clean_texts.py
│   ├── linguistic_features.py
│   ├── analyze_composition_mfmd.py
│   ├── build_mfmd_workbook.mjs
│   └── build_mfmd_report.py
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
    └── mfmd_analysis/
        ├── figures/
        ├── tables/
        ├── analysis_metadata.json
        ├── workbook_payload.json
        └── 作文多维分析结果.xlsx
```

### 核心数据文件

- `4.14作文前筛.xlsx`：原始筛选表，包含多个 sheet，是后续抽样的源数据。
- `作文样本主表.xlsx`：当前主要迭代用表。四个 sheet 分别为 `J1`、`J2`、`Y1`、`Y2`，每个 sheet 155 篇。表内包含作文基本信息，并增加了 `作文文件名` 字段，例如 `J1_55_01`、`J1_60_02`。
- `outputs/新版HSK词汇大纲.csv`：从新版 HSK 考试大纲提取的 11,000 条机器可读词汇，包含主等级、兼属等级、词语、拼音和词性等字段。
- `作文词性统计宽表.xlsx`：项目最终主分析表，基于 `clean_text` 全量生成。`词性统计` sheet 每行对应一篇作文，`字段说明` sheet 逐列记录定义、公式和分母；后续统计分析应优先使用此表。
- `作文语言特征多维分析报告.docx`、`作文语言特征多维分析报告.pdf`：基于620篇作文和301列主宽表生成的29页中文MF/MD学术分析报告，包含11幅统计图、五维解释、四组比较、分数关系、补充国籍分析及匿名文本例证。
- `outputs/mfmd_analysis/作文多维分析结果.xlsx`：MF/MD分析的可审查结果工作簿，共20个sheet，保存变量筛选、候选模型诊断、因子载荷、维度得分、组间检验、稳健回归、文本例证和图表索引。
- `outputs/mfmd_analysis/tables/`、`outputs/mfmd_analysis/figures/`：分析脚本生成的逐表CSV和11幅300 DPI图片；`analysis_metadata.json`保存随机种子、诊断和最终模型元数据。
- `outputs/hsk_标注说明.md`：从 HSK 网站帮助页整理出的标注说明，用于理解和清洗原始标注文本。

### 文本目录

- `ori_text/{篇名代码}/{作文编码}.txt`：从 HSK 接口抓取的原始标注文本，保留网站返回的标注内容。
- `clean_text/{篇名代码}/{作文文件名}.txt`：清洗后的正文文本，只保留可确定的正确文本，保留原文目录不变。
- `seg_text/{篇名代码}/{作文文件名}.txt`：PyNLPIR 分词和词性标注结果。每行格式为 `词/中文词性 词/中文词性 ...`，保留原段落换行和标点。

当前三个文本目录均为 620 个 `.txt` 文件，其中 `J1/J2/Y1/Y2` 各 155 篇。

### 脚本

- `scripts/sample_essays_by_score.py`：按分数段抽样作文，尽量保持国籍多样和数量均衡，并排除中国香港、中国台湾、中国澳门和中国少数民族样本。
- `scripts/fetch_hsk_ori_texts.py`：批量从 HSK 网站接口抓取原始标注文本，保存到 `ori_text`。
- `scripts/clean_hsk_ori_texts.py`：清洗 `ori_text`，生成 `clean_text`，不修改原始文本。
- `scripts/extract_hsk_vocabulary.py`：从新版 HSK 考试大纲 PDF 提取机器可读词汇 CSV。
- `scripts/segment_hsk_clean_texts.py`：使用 PyNLPIR 对 `clean_text` 分词，生成 `seg_text` 和语言特征宽表；细粒度词性用于语法特征，大类中文词性继续写入分词文本。
- `scripts/linguistic_features.py`：不依赖 PyNLPIR 运行环境的纯计算模块，负责词汇、句段、语法、篇章、记叙描写和 HSK 派生特征。
- `scripts/analyze_composition_mfmd.py`：以根目录主宽表为输入，完成变量筛选、平行分析、MinRes/Promax因子分析、Bootstrap稳定性、组间比较和稳健回归，并输出CSV、JSON与统计图。
- `scripts/build_mfmd_workbook.mjs`：使用 `@oai/artifact-tool` 将分析JSON整理为多sheet结果工作簿。
- `scripts/build_mfmd_report.py`：根据正式分析结果和图片生成中文DOCX分析报告；PDF由最终DOCX统一导出。
- `requirements-analysis.txt`：MF/MD分析与报告的固定版本Python依赖。
- `resources/语言特征词表.csv`：可人工审查的语言特征词表，字段为 `大类、特征名、词项、允许词性前缀、来源说明`。
- `tests/test_linguistic_features.py`：纯函数单元测试，覆盖 MATTR、句段切分、词表最长匹配、连续动词和 HSK 派生统计等口径。

## 常用运行命令

抓取原始文本：

```bash
python scripts/fetch_hsk_ori_texts.py
```

清洗原始文本：

```bash
python scripts/clean_hsk_ori_texts.py
```

分词并生成词性统计宽表：

```bash
arch -x86_64 /usr/bin/python3 scripts/segment_hsk_clean_texts.py
```

该命令默认在项目根目录生成或更新最终主分析表 `作文词性统计宽表.xlsx`。

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

## MF/MD 多维分析

正式分析固定随机种子为 `20260829`。脚本从146项标准化候选指标出发，经方差/稀疏性、相关性、变量级MSA和共同度筛选，最终保留39项指标并提取5个维度：

1. 词汇丰富度与词汇扩展
2. 基础词汇与叙事推进
3. 人称指涉与信息密度
4. 句法延展与分句复杂度
5. 动作过程与动词链

最终模型 `KMO=0.752`，累计解释方差为 `57.3%`；200次Bootstrap全部成功，各维度Tucker一致性系数中位数均不低于 `0.992`。维度载荷正负表示语言指标的共现方向，不代表作文质量高低；`J1/J2` 与 `Y1/Y2` 的差异只解释为题目与体裁组合差异。

复现分析：

```bash
python -m venv .venv-analysis
source .venv-analysis/bin/activate
pip install -r requirements-analysis.txt
python scripts/analyze_composition_mfmd.py
python scripts/build_mfmd_report.py
```

若运行环境已提供 `@oai/artifact-tool`，可再生成结果工作簿：

```bash
node scripts/build_mfmd_workbook.mjs
```

分析脚本默认使用1000次Horn平行分析和200次Bootstrap；输出写入 `outputs/mfmd_analysis/`，不修改主宽表、两篇论文或文本语料。

## 当前数据流

```text
4.14作文前筛.xlsx
  -> scripts/sample_essays_by_score.py
  -> 作文样本主表.xlsx
  -> scripts/fetch_hsk_ori_texts.py
  -> ori_text/
  -> scripts/clean_hsk_ori_texts.py
  -> clean_text/
  -> scripts/segment_hsk_clean_texts.py
       + outputs/新版HSK词汇大纲.csv
       + resources/语言特征词表.csv
       + scripts/linguistic_features.py
  -> seg_text/
  -> 作文词性统计宽表.xlsx
  -> scripts/analyze_composition_mfmd.py
  -> outputs/mfmd_analysis/{tables,figures,analysis_metadata.json,workbook_payload.json}
  -> outputs/mfmd_analysis/作文多维分析结果.xlsx
  -> 作文语言特征多维分析报告.{docx,pdf}
```

`作文样本主表.xlsx` 是样本主索引；`clean_text` 是文本分析的正文来源；根目录的 `作文词性统计宽表.xlsx` 是项目最终主分析表，`seg_text` 是其可复核的分词中间结果。根目录的DOCX/PDF是当前正式分析报告，详细统计表位于 `outputs/mfmd_analysis/作文多维分析结果.xlsx`。
