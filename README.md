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
├── 母语作文样本主表.xlsx
├── 母语作文词性统计宽表.xlsx
├── 作文语言特征母语对照分析报告.docx
├── 作文语言特征母语对照分析报告.pdf
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
│   ├── build_mfmd_report.py
│   ├── collect_zuowen_native_controls.py
│   ├── segment_native_control_texts.py
│   ├── build_native_control_workbooks.mjs
│   ├── analyze_native_control.py
│   ├── build_native_control_analysis_workbook.mjs
│   └── build_native_control_report.py
├── resources/
│   ├── 语言特征词表.csv
│   └── native_control_manual_review.csv
├── tests/
│   ├── fixtures/
│   ├── test_linguistic_features.py
│   └── test_zuowen_native_controls.py
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
├── native_ori_text/        # 本地保存，Git忽略
├── native_clean_text/      # 本地保存，Git忽略
├── native_seg_text/        # 本地保存，Git忽略
└── outputs/
    ├── mfmd_analysis/
        ├── figures/
        ├── tables/
        ├── analysis_metadata.json
        ├── workbook_payload.json
        └── 作文多维分析结果.xlsx
    ├── native_control/
    │   ├── selected_samples.json
    │   ├── selection_summary.json
    │   ├── candidate_audit.csv
    │   └── native_stats_payload.json
    └── native_control_analysis/
        ├── figures/
        ├── tables/
        ├── analysis_metadata.json
        ├── workbook_payload.json
        └── 作文母语对照分析结果.xlsx
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
- `母语作文样本主表.xlsx`：作文网高中作文参照样本的来源审计主表，记录 `NJ1/NJ2/NY1/NY2`、对应学习者题目、URL、发布日期、年级、主题匹配层级、篇幅偏差、正文哈希和审核状态；不收录完整网页正文。
- `母语作文词性统计宽表.xlsx`：180篇母语参照语料的301列派生统计，字段顺序和口径与学习者主宽表一致；作文分数留空，国籍统一标记为“中国（公开网络样本）”。
- `作文语言特征母语对照分析报告.docx`、`作文语言特征母语对照分析报告.pdf`：29页详细比较报告，将母语参照作文投影到现有五维学习者量尺，包含采样审计、篇幅匹配、Welch检验、Hedges `g`、HC3回归、重抽样、主题/年份敏感性和联合因子分析。
- `outputs/native_control_analysis/作文母语对照分析结果.xlsx`：母语对照分析的可审查结果工作簿，保存投影参数、逐篇得分、组间检验、39项特征比较、稳健回归、敏感性和联合模型结果。

### 文本目录

- `ori_text/{篇名代码}/{作文编码}.txt`：从 HSK 接口抓取的原始标注文本，保留网站返回的标注内容。
- `clean_text/{篇名代码}/{作文文件名}.txt`：清洗后的正文文本，只保留可确定的正确文本，保留原文目录不变。
- `seg_text/{篇名代码}/{作文文件名}.txt`：PyNLPIR 分词和词性标注结果。每行格式为 `词/中文词性 词/中文词性 ...`，保留原段落换行和标点。

当前三个文本目录均为 620 个 `.txt` 文件，其中 `J1/J2/Y1/Y2` 各 155 篇。

母语参照全文目录 `native_ori_text`、`native_clean_text`、`native_seg_text` 以及网页缓存 `native_cache` 均只保存在本地并排除 Git。当前三个母语文本目录均为180个 `.txt`，其中 `NJ1/NJ2/NY1/NY2` 各45篇，最终样本全部来自高中范围。作文网高中栏目能够证明文章被网站归入高中作文，但不能独立核验作者身份，也不能完全排除编辑、转载或润色，因此本项目统一称其为“公开网络母语参照语料”，不把它表述为身份认证的母语者实验语料。

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
- `scripts/collect_zuowen_native_controls.py`：低速单线程发现、抓取、解析和筛选作文网高中作文，支持缓存断点续跑、指数退避、主题分级、篇幅匹配、哈希和字符五元组去重，并输出完整候选审计。
- `scripts/segment_native_control_texts.py`：复用现有PyNLPIR和七类语言特征逻辑，对母语清洗文本分词并生成与学习者宽表对齐的301列统计载荷。
- `scripts/build_native_control_workbooks.mjs`：使用 `@oai/artifact-tool` 生成母语样本主表和母语词性统计宽表。
- `scripts/analyze_native_control.py`：固定使用620篇学习者作文的39项指标投影参数和五维结构，执行母语对照、HC3回归、篇幅匹配重抽样、主题/年份敏感性和联合EFA。
- `scripts/build_native_control_analysis_workbook.mjs`：生成多sheet母语对照分析结果工作簿。
- `scripts/build_native_control_report.py`：根据正式统计表和300 DPI图片生成母语对照DOCX报告；PDF由最终DOCX统一导出。
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

运行测试：

```bash
python -m unittest discover -s tests -v
```

## 公开网络母语参照语料

目标设计为 `NJ1/NJ2/NY1/NY2` 每组50篇。穷尽允许的高中来源并完成跨组去重后，`NJ2` 只有45篇，因此最终统一为每组45篇、共180篇。完整高中单元作文和练习来源作文可以纳入；只有题目、要求、讲解或素材的页面仍排除。脚本为 `NY2` 保留初中作文最后备用通道，但本轮最终样本没有使用初中来源，也没有使用站外来源。

主题匹配分为 `精确`、`近似`、`扩展`、`宽泛` 四级。`NJ2` 的宽泛层扩展到个人经历，`NY1` 扩展到健康、公共责任、规则及相邻社会议题，`NY2` 扩展到家庭影响、教育、自立和青少年成长。当前精确或近似主题样本量仅为 `J1=9`、`J2=8`、`Y1=2`、`Y2=4`，所以后三组主要是题目域参照，不能表述为严格同题对照。

网页作文整体长于学习者作文。1000次最近邻篇幅匹配后，母语参照仍比对应学习者平均多约 `40/119/192/79` 个汉字（J1/J2/Y1/Y2）。报告使用 `log(纯文本字数)` 的HC3稳健回归和篇幅重抽样降低影响，但无法彻底消除篇幅混杂；完整限制和敏感性结果见正式报告。网页全文仅用于本地研究；作文网服务协议对内容复制与传播设有限制，因此公开仓库只保存脚本、URL、元数据、派生统计和报告。

首次采集或从缓存续跑：

```bash
python3 scripts/collect_zuowen_native_controls.py
```

PyNLPIR分词与301项特征统计：

```bash
arch -x86_64 /usr/bin/python3 scripts/segment_native_control_texts.py
```

生成两份母语工作簿：

```bash
node scripts/build_native_control_workbooks.mjs
```

分析和报告环境要求Python 3.10以上，当前使用Python 3.12：

```bash
python3 -m venv .venv-analysis
source .venv-analysis/bin/activate
pip install -r requirements-analysis.txt
python scripts/analyze_native_control.py
node scripts/build_native_control_analysis_workbook.mjs
python scripts/build_native_control_report.py
```

母语对照主分析沿用当前39项指标形成的五维结构，以学习者样本均值、标准差和载荷方向固定投影；联合因子分析仅作为结构敏感性检验。母语样本没有作文分数，因此不进行母语与学习者的分数比较。当前800篇联合分析中，只有2因子方案通过预设诊断和稳定性门槛，不能视为对原五维结构的完整复现。

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

作文网高中作文栏目
  -> scripts/collect_zuowen_native_controls.py
  -> native_ori_text/ + native_clean_text/（本地、Git忽略）
  -> 母语作文样本主表.xlsx
  -> scripts/segment_native_control_texts.py
  -> native_seg_text/（本地、Git忽略）
  -> 母语作文词性统计宽表.xlsx
  -> scripts/analyze_native_control.py
  -> outputs/native_control_analysis/{tables,figures,作文母语对照分析结果.xlsx}
  -> 作文语言特征母语对照分析报告.{docx,pdf}
```

`作文样本主表.xlsx` 是样本主索引；`clean_text` 是文本分析的正文来源；根目录的 `作文词性统计宽表.xlsx` 是项目最终主分析表，`seg_text` 是其可复核的分词中间结果。根目录的DOCX/PDF是当前正式分析报告，详细统计表位于 `outputs/mfmd_analysis/作文多维分析结果.xlsx`。
