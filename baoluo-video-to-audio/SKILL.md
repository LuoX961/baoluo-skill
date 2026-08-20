---
name: baoluo-video-to-audio
description: 视频转音频，使用 ffmpeg 超快速提取音频为 M4A 格式，直接复制音频流
---

# 视频转音频 Skill

> 负反馈清单：[feedback/规则索引.csv](feedback/规则索引.csv)

## 类型

无状态型：提取音频写入输出目录，不读取历史。单次调用自我闭合，不设治理机制。

从视频文件中超快速提取音频，输出为 M4A 格式。

## 功能特点

- 使用 ffmpeg 直接复制音频流，不重新编码
- 输出 M4A 格式（与视频原始音频格式一致）
- 处理速度极快，1 小时视频约 2-3 秒完成（3000 倍速）

## 支持格式

**视频格式**：MP4, MOV, AVI, MKV, FLV, WMV, WEBM, M4V 等

## 执行流程

1. 询问用户视频文件路径和输出目录
2. 使用 ffmpeg 直接复制音频流（不重新编码）
3. 输出 M4A 音频文件

## 使用方法

触发方式：
- `$baoluo-video-to-audio`
- 或直接说「把这个视频转成音频」

## 执行步骤

### 1. 获取文件信息

询问用户：
- 视频文件路径（必须）
- 输出目录（可选，默认为视频所在目录）

### 2. 执行转换

调用脚本：`<Skill 目录>/video_to_audio_fast.py`

命令格式：
```bash
python3 "<Skill 目录>/video_to_audio_fast.py" "<视频路径>" "<输出目录>"
```

**脚本特性**：
- 使用 `ffmpeg -c:a copy` 直接复制音频流
- 不重新编码，速度极快（约 3000 倍速）
- 保持原始音频质量

### 3. 输出结果

- 生成与源文件同名的 .m4a 文件
- 告知用户输出文件位置

## 错误处理

- 文件不存在：提示用户检查路径
- 格式不支持：列出支持的格式
- 未安装 ffmpeg：提示执行 `brew install ffmpeg`

## 脚本维护

`<Skill 目录>` 指安装后 `baoluo-video-to-audio` 文件夹的实际位置。

如需修改输出格式、添加其他参数等，直接编辑该脚本文件。

## 示例

**用户输入**：
```
$baoluo-video-to-audio
文件：/path/to/video/demo.mp4
输出：/path/to/audio
```

**执行结果**：
```
📹 输入文件：demo.mp4
🎵 输出文件：/path/to/audio/demo.m4a
⏳ 正在提取音频...
✅ 提取成功：/path/to/audio/demo.m4a
```
