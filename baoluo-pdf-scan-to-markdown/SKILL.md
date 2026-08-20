---
name: baoluo-pdf-scan-to-markdown
description: 将扫描件／图片型 PDF 转换为干净的 Markdown（正文 + 章节标题），自动清除页眉、页脚、页码、版权页、扉页和目录，识别章节标题，处理双栏排版，合并跨页段落。全程本地 macOS 原生 Vision OCR，不消耗 LLM Token。当用户说「这个 PDF 是扫描件」「扫描版转文字」「图片型 PDF 转 Markdown」「PDF 转 md」「识别这本书」「转扫描书」或提供 .pdf 扫描文件路径时使用。不适合文字型（可选中复制）PDF——那种直接 pdftotext 或 pandoc 更快。
---

# PDF 扫描件 → Markdown 转换器

> 负反馈清单：[feedback/规则索引.csv](feedback/规则索引.csv)

## 类型

无状态型：OCR 转换后写入目标目录，不读取以往转换结果。单次调用自我闭合，不设治理机制。

把扫描件 / 图片型 PDF（文字不可选中）转成干净、可读、只有正文和章节标题的 Markdown。

## 核心特性

| 能力 | 说明 |
|---|---|
| OCR 引擎 | macOS 原生 Vision 框架（本地、免费、中文识别质量优） |
| Token 消耗 | 零（OCR 不经过 LLM） |
| 页眉/页脚/页码 | 按坐标区域自动清除 |
| 章节标题 | 自动识别（含页眉型章首与独立大标题页），转 `# 标题` |
| 双栏排版 | 按文字 x 坐标拆左右栏重组，跨栏段落可复原 |
| 封面/扉页/版权页/目录 | 自动剔除 |
| 段落合并 | 跨页段落自动续接 |

## 依赖

- macOS 12+（Vision 框架）
- Python 3（Homebrew 版即可）
- pyobjc-framework-Vision、pyobjc-framework-Quartz（由 setup.sh 安装到 `~/.venvs/ocr`）

## 安装（每台电脑一次）

```bash
bash "<Skill 目录>/setup.sh"
```

## 执行流程（两步）

```bash
# 第 1 步：OCR 识别，输出带坐标的 JSON（速度约 1 秒/页）
"${HOME}/.venvs/ocr/bin/python" "<Skill 目录>/ocr_vision.py" \
  "/路径/到/书.pdf" \
  "/临时目录/book_ocr.json"

# 第 2 步：清理转换，输出 Markdown（输出位置由用户指定）
"${HOME}/.venvs/ocr/bin/python" "<Skill 目录>/clean_ocr.py" \
  "/临时目录/book_ocr.json" \
  "/输出位置/书名.md"
```

`<Skill 目录>` 指安装后 `baoluo-pdf-scan-to-markdown` 文件夹的实际位置。

## 第 2 步完成后

- 脚本会打印「标题数」和全部章节标题，用于校验
- 抽查正文开头、中间、结尾三段，确认无残留页眉页码
- 交付前向用户说明 OCR 局限（见下）

## 已知局限与处理

| 情况 | 处理 |
|---|---|
| OCR 错字（自由→白由、赢→嬴、乐部→俱乐部 等） | 标题级自动纠错已内置；正文保留原样，用户可自行校对 |
| 行首漏字（双栏页常见） | 无法自动恢复，双栏重组后顺序已正确 |
| 新排版形态（页眉位置/字号不同） | 需调整 `clean_ocr.py` 中的阈值（页眉区 y>0.88、窄行 w<0.35、字号 h≥0.024）并重跑第 2 步 |
| 公式、复杂表格、手写体 | Vision 支持有限，识别质量无法保证 |
| 竖排、倾斜页 | 可能识别错乱，建议先人工修正页面方向 |

## 注意事项

- 转换后**不删除原始 PDF**，除非用户明确要求
- 输出位置由用户指定；若存入 Obsidian 知识库，常见位置是 `00-常用&收件箱/01-书籍/《书名》/`
- 中间 JSON 文件（数百 KB～数 MB）建议保留到用户确认成果后再清理，便于重新调整清理参数
- 跨设备：依赖装好后两台电脑均可用；本 Skill 不在单台电脑维护独立副本
