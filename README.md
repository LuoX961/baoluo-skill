# baoluo-skill

这是 baoluo 维护的 Agent Skill 合集。每个子目录都是一个可以独立安装的 Skill。

## Skill 列表

| Skill | 用途 | 主要依赖 |
|---|---|---|
| `baoluo-pdf-scan-to-markdown` | 使用 macOS Vision OCR 将扫描版 PDF 转换为 Markdown | macOS 12+、Python 3、PyObjC |
| `baoluo-qianji-bill-cleaner` | 按固定规则清理钱迹导出的 CSV 或 XLSX 账单 | Node.js、支持 `@oai/artifact-tool` 的电子表格运行环境 |
| `baoluo-video-to-audio` | 使用 FFmpeg 从视频中提取 M4A 音频 | Python 3、FFmpeg |

## 安装

下载整个仓库：

```bash
git clone https://github.com/LuoX961/baoluo-skill.git
```

从仓库中选择需要的 Skill 文件夹，将整个文件夹复制到所用 Agent 平台的 Skill 目录。不要只复制 `SKILL.md`，脚本和其他资源也属于 Skill 的一部分。

以 Codex 为例，可以将需要的文件夹复制到 `~/.codex/skills/`。安装后重启或刷新 Agent，使新 Skill 被重新发现。

每个 Skill 的具体依赖和使用方法以对应目录中的 `SKILL.md` 为准。

## 许可证

本仓库使用 [MIT License](LICENSE)。
