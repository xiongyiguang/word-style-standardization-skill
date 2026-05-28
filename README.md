# word-style-standardization

该项目用于将 Word `.docx` 文件按标准模板进行样式规整，重点采用“原位处理”方式，避免图片、表格、媒体、页眉页脚和对象关系丢失。

适用于投标文件、方案文档、咨询报告、质量管理材料等中文 Word 文档的批量样式统一。

## 核心原则

- 不抽取正文文本重建 Word。
- 保留源 `.docx` 包结构、媒体、图片关系、表格、页眉页脚、对象锚点和嵌入对象。
- 只把源文档已有 heading 1-9 语义的段落映射为模板标题样式，不凭文本编号猜标题。
- 正文里的 `一、二、三`、`1、2、3`、`1）、2）、3）` 保留为手工编号文本，不套模板编号样式。
- 正文自动编号会先转成段首手工编号文本，再移除 `w:numPr`。
- 只有文本自带明确项目符号或自动编号定义为 bullet 时才套箭头/打钩/四角星列表样式；数字编号不会误套箭头，避免出现“箭头 + 1)”的重复编号。
- 表格本体会套用模板中的“表格标准样式”，表格内段落会套用 `P_普通图表标准格式_表格`。
- 清理会覆盖模板样式的直接格式：表格单元格边框/底纹、标题直接格式、正文基础字体字号颜色和段落间距缩进。
- 生成规整后的 `.docx` 和校验报告。

## 目录结构

```text
word-style-standardization-skill/
├── SKILL.md
├── README.md
├── requirements.txt
├── scripts/
│   └── normalize_word_style.py
├── references/
│   └── checklist.md
├── dist/
│   └── claude-word-style-standardization.zip
└── assets/
    └── 标准样式Word文档.docx
```

## 使用示例

```bash
python -m venv .venv
.venv/bin/pip install -r requirements.txt
mkdir -p input output
.venv/bin/python scripts/normalize_word_style.py \
  --input input/source.docx \
  --template assets/标准样式Word文档.docx
```

默认输出到 `output/<源文件名>_标准样式规整.docx`，校验报告输出到 `output/<源文件名>_样式规整校验报告.md`。如需自定义路径，可显式传入 `--output` 和 `--report`。

## 给同事的提示词示例

请使用 `word-style-standardization` 技能，将 `input/source.docx` 按 `assets/标准样式Word文档.docx` 进行样式规整。要求采用原位处理方式，不要重建正文；保留图片、表格、页眉页脚、对象锚点、媒体引用和嵌入对象；生成规整后的 docx，并输出校验报告。

## Claude 版本

Claude 可用版本已打包在：

```text
dist/claude-word-style-standardization.zip
```
