# docx_parsing
docx parsing with xml
# DOCX 文档解析工具

基于 XML 解析思路的 DOCX 文档内容提取工具，直接解析 docx 压缩包内的 XML 文件，无需依赖 Microsoft Office 或第三方 docx 库。

## 功能特性

- **文本提取**：精确提取段落文本，保留样式信息（对齐、缩进、间距等）
- **表格提取**：保持完整的行列结构，提取单元格内容和属性（合并、宽度等）
- **标题编号提取**：支持自动编号和手动编号两种来源
  - 自动编号：通过 `numbering.xml` 编号定义 + `styles.xml` 关联计算
  - 手动编号：从标题文本中正则提取（如 `6.8.1 供应国资格` → `6.8.1`）
  - 支持正向关联（`styles.xml` 中的 `numId`）和反向关联（`numbering.xml` 中的 `pStyle`）
- **标题层级提取**：识别标题样式，构建文档大纲
- **审阅内容提取**：
  - 修订标记：插入（`w:ins`）和删除（`w:del`）
  - 评论：从 `comments.xml` 提取评论内容和引用关系
  - 支持四种处理模式：`final` / `original` / `both` / `track`
- **图片提取与 OCR**：提取图片文件，可选集成 Tesseract OCR 进行文字识别
- **编号格式支持**：decimal、字母、罗马数字、中文数字等多种编号格式

## 依赖

```
pip install beautifulsoup4
```

OCR 功能（可选）：

```
pip install pytesseract Pillow
```

## 使用方式

### 命令行

```bash
# 基本用法 - 输出 JSON
python docx_processor.py input.docx -o ./output

# 输出纯文本格式
python docx_processor.py input.docx -o ./output --export text

# 启用 OCR 图片识别
python docx_processor.py input.docx -o ./output --ocr

# 指定审阅处理模式
python docx_processor.py input.docx -o ./output --review-mode both
```

### 命令行参数

| 参数 | 说明 | 默认值 |
|------|------|--------|
| `input` | 输入 docx 文件路径 | 必填 |
| `-o, --output` | 输出目录 | `./output` |
| `--ocr` | 启用 OCR 图片识别 | 关闭 |
| `--review-mode` | 审阅处理模式：`final` / `original` / `both` / `track` | `final` |
| `--export` | 导出格式：`json` / `text` | `json` |

### Python API

```python
from docx_processor import EnterpriseDocxProcessor

processor = EnterpriseDocxProcessor(
    file_path='input.docx',
    output_dir='./output',
    enable_ocr=False,
    review_mode='final'
)

result = processor.extract_comprehensive_content()
```

## 输出结构

```json
{
  "metadata": {
    "processing_mode": "final",
    "ocr_enabled": false,
    "extraction_timestamp": "2026-06-02T10:00:00"
  },
  "content": {
    "structured_elements": [
      {
        "type": "paragraph",
        "content": "段落文本",
        "style_info": { "style_id": "1", "alignment": "center" }
      },
      {
        "type": "table",
        "rows": [[{ "content": "单元格", "properties": {} }]],
        "properties": { "style": "TableGrid" }
      }
    ],
    "heading_outline": [
      { "level": 1, "text": "标题文本", "number": "6.8", "style_id": "21", "style_name": "heading 2" }
    ],
    "review_summary": {
      "total_changes": 3,
      "insertions": 1,
      "deletions": 0,
      "comments": 5,
      "details": [...]
    },
    "images": { "rId1": "OCR识别文本" }
  }
}
```

## 核心类

### EnterpriseDocxProcessor

主处理器，负责协调各提取模块并合并结果。

| 方法 | 说明 |
|------|------|
| `extract_comprehensive_content()` | 综合内容提取主入口 |
| `extract_basic_elements(archive)` | 提取段落、表格、图片 |
| `extract_heading_numbering(archive)` | 提取标题编号系统 |
| `extract_heading_hierarchy(archive)` | 提取标题层级结构 |
| `process_review_changes(archive)` | 处理审阅修改和评论 |
| `merge_extraction_results(results)` | 整合所有提取结果 |

### ReviewContentProcessor

审阅内容处理器，支持修订标记和评论的提取。

| 方法 | 说明 |
|------|------|
| `process_document_reviews(document_xml, comments_xml)` | 处理所有审阅内容 |
| `process_paragraph_reviews(paragraph, comments_map)` | 处理单个段落的审阅标记 |
| `extract_collaboration_history(archive)` | 提取协作历史（作者、时间线、评论线程） |

## 审阅处理模式

| 模式 | 说明 |
|------|------|
| `final` | 只保留最终文本（插入内容），忽略删除内容 |
| `original` | 只保留原始文本（删除内容），忽略插入内容 |
| `both` | 同时保留原始和修改后的文本 |
| `track` | 保留完整的修改轨迹 |

## 编号提取策略

标题编号提取采用两级策略：

1. **自动编号（优先）**：当段落的样式在编号映射中时，根据 `numbering.xml` 中的编号定义和计数器计算编号字符串
2. **文本正则提取（回退）**：当标题样式没有关联自动编号时，从标题文本中用正则提取编号

样式与编号的关联支持两种路径：
- **正向关联**：`styles.xml` 中样式定义的 `w:pPr/w:numId` 直接引用编号 ID
- **反向关联**：`numbering.xml` 中 `w:lvl/w:pStyle` 引用样式 ID（飞书等导出的 docx 常用此方式）

支持的文本编号格式：
- `6.8 标题` / `6.8.1 标题` → 多级数字编号
- `1. 标题` → 带点号的数字编号
- `（1）标题` / `(1) 标题` → 括号编号
- `第一章标题` / `第二节标题` → 中文章节编号

## 注意事项

- OCR 功能需要安装 [Tesseract OCR](https://github.com/tesseract-ocr/tesseract) 并配置中文语言包（`chi_sim`）
- EMF/WMF 格式的图片暂不支持自动转换，会保留原路径并输出警告
- 文档中的修订标记需处于"显示修订"状态才能被提取
- `numbering.xml` 不是所有 docx 文件都包含，缺失时标题编号提取将回退到文本正则方式
