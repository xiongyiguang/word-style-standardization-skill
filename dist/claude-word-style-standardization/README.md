# word-style-standardization Claude Skill

这个 Claude Skill 用于将 `.docx` 文件按标准 Word 模板进行原位样式规整。

重点原则：

- 不抽取文本重建 Word。
- 保留源 `.docx` 包结构、图片、表格、页眉页脚、对象锚点、媒体引用和嵌入对象。
- 只把源文档已有 heading 1-9 语义的段落映射为模板标题样式。
- 正文中的 `一、二、三`、`1、2、3`、`1）、2）、3）` 保留为手工编号文本，不套模板编号样式。
- 正文自动编号会先转成段首手工编号文本，再移除 `w:numPr`。
- 只有文本自带明确项目符号或自动编号定义为 bullet 时才套箭头/打钩/四角星列表样式；数字编号不会误套箭头，避免出现“箭头 + 1)”的重复编号。
- 输出规整后的 `.docx` 和校验报告。

## 目录

```text
claude-word-style-standardization/
├── SKILL.md
├── README.md
├── requirements.txt
├── scripts/
│   └── normalize_word_style.py
├── references/
│   └── checklist.md
└── assets/
    └── 标准样式Word文档.docx
```

## 使用

```bash
python -m venv .venv
.venv/bin/pip install -r requirements.txt
mkdir -p input output
.venv/bin/python scripts/normalize_word_style.py \
  --input input/source.docx \
  --template assets/标准样式Word文档.docx
```

默认输出到 `output/<源文件名>_标准样式规整.docx`，校验报告输出到 `output/<源文件名>_样式规整校验报告.md`。如需自定义路径，可显式传入 `--output` 和 `--report`。

## 给使用者的提示词

请使用 `word-style-standardization` 技能，将 `input/source.docx` 按 `assets/标准样式Word文档.docx` 进行样式规整。要求采用原位处理方式，不要重建正文；保留图片、表格、页眉页脚、对象锚点、媒体引用和嵌入对象；不要凭文本推断标题；正文自动编号转为手工编号；生成规整后的 docx，并输出校验报告。
