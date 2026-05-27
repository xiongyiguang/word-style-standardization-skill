---
name: word-style-standardization
description: Standardize Chinese .docx documents against the bundled standard Word template by modifying styles in-place, preserving media, tables, headers, footers, relationships, object anchors, embedded objects, and producing a verification report.
---

# Word Style Standardization

Use this skill when the user asks to standardize, normalize, clean up, or apply a standard Word style template to a `.docx` document, especially Chinese tender, proposal, solution, quality-management, or report documents.

## Core Rule

Do not rebuild the Word document from extracted text.

Always preserve the original `.docx` package structure and relationships. The script copies the source document, imports style-related parts from the template, and edits paragraph style properties in WordprocessingML.

This avoids losing images, tables, headers, footers, object anchors, media references, embedded objects, section structure, comments, footnotes, endnotes, bookmarks, fields, and cross-references.

## Bundled Files

- `assets/标准样式Word文档.docx`: standard Word style template
- `scripts/normalize_word_style.py`: in-place `.docx` style normalization script
- `references/checklist.md`: acceptance checklist
- `requirements.txt`: Python dependency list

## Recommended Command

```bash
python scripts/normalize_word_style.py \
  --input input/source.docx \
  --template assets/标准样式Word文档.docx \
  --output output/source_标准样式规整.docx \
  --report output/source_样式规整校验报告.md
```

If Python dependencies are missing, install them in a virtual environment:

```bash
python -m venv .venv
.venv/bin/pip install -r requirements.txt
.venv/bin/python scripts/normalize_word_style.py \
  --input input/source.docx \
  --template assets/标准样式Word文档.docx \
  --output output/source_标准样式规整.docx \
  --report output/source_样式规整校验报告.md
```

## Style Rules

The script applies the bundled template styles:

| Content | Template styleId |
| --- | --- |
| Heading 1 | `10` |
| Heading 2 | `20` |
| Heading 3 | `30` |
| Heading 4 | `40` |
| Heading 5 | `50` |
| Heading 6 | `6` |
| Heading 7 | `7` |
| Heading 8 | `8` |
| Heading 9 | `9` |
| Body paragraph | `P1505` |
| Short bold body paragraph | `P` |
| Table paragraph | `P5` |
| Image/object paragraph | `P6` |
| Explicit bullet list level 1 | `P2` |
| Explicit bullet list level 2 | `P3` |
| Explicit bullet list level 3 | `P4` |

## Important Boundaries

- Only apply heading styles when the source paragraph already has heading 1-9 semantics, or its source style inherits from heading 1-9.
- Do not infer headings from text such as `1 xxx`, `1.1 xxx`, `一、xxx`, or numeric formulas.
- Body text that starts with `一、二、三`, `1、2、3`, or `1）、2）、3）` must remain manual numbering text. Do not apply template numbering styles such as `P1`, `P2123`, or `P3123`.
- For non-heading body paragraphs with Word automatic numbering, first compute the visible numbering text from the source `word/numbering.xml`, insert it at the start of the paragraph as normal text, and then remove `w:numPr`.
- Do not manually add numbering text to headings; heading numbering should remain controlled by the template heading styles.
- Use explicit bullet list styles only when the paragraph text itself starts with clear bullet markers such as `•`, `●`, `▪`, `→`, `✓`, or dash bullets, or when the source automatic numbering definition has `numFmt=bullet`.
- Do not apply `P2/P3/P4` to decimal, Chinese, or parenthesized automatic numbering. Convert those to manual text such as `1)`, `1、`, or `一、`, then use a body style; otherwise the output can show duplicated markers such as an arrow plus `1)`.

## Required Verification

After generating the output, inspect the report and confirm:

- Media file count did not decrease.
- Image relationship count did not decrease.
- Paragraph count is unchanged.
- Table count is unchanged.
- Normal/body-default style usage is reduced.
- Non-heading automatic numbering has been removed or reduced after manual-number conversion.
- Heading style statistics match source heading semantics.

If media files, image relationships, or tables decrease, treat the output as failed and do not deliver it.

## Delivery Response

When finished, provide:

- the normalized `.docx` path
- the verification report path
- a brief note that the process was in-place and preserved document relationships
