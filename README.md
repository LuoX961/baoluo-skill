# baoluo-skill

这是 baoluo 维护的 Agent Skill 合集。每个子目录都是一个可以独立安装的 Skill。

## Skill 列表

| Skill | 用途 | 主要依赖 |
|---|---|---|
| `baoluo-pdf-scan-to-markdown` | 使用 macOS Vision OCR 将扫描版 PDF 转换为 Markdown | macOS 12+、Python 3、PyObjC |
| `baoluo-qianji-bill-cleaner` | 按日期清理钱迹 CSV／XLSX 账单，备份并校验后安全更新原财务统计表 | Python 3（仅使用标准库） |
| `baoluo-video-to-audio` | 使用 FFmpeg 从视频中提取 M4A 音频 | Python 3、FFmpeg |

## 安装

只安装钱迹账单清理 Skill：

```bash
npx -y skills add LuoX961/baoluo-skill -g --skill baoluo-qianji-bill-cleaner --agent '*'
```

安装仓库中的全部 Skill：

```bash
npx -y skills add LuoX961/baoluo-skill -g --all
```

也可以下载整个仓库后手动安装：

```bash
git clone https://github.com/LuoX961/baoluo-skill.git
```

手动安装时，将需要的完整 Skill 文件夹复制到 Agent 平台的 Skill 目录。不要只复制 `SKILL.md`，脚本和其他资源也属于 Skill 的一部分。

每个 Skill 的具体依赖、输入格式和安全边界以对应目录中的 `SKILL.md` 为准。

## 许可证

本仓库使用 [MIT License](LICENSE)。
