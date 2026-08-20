#!/usr/bin/env python3
"""
视频转音频 - 超快速版
直接复制音频流，不重新编码，速度极快
"""

import sys
import os
from pathlib import Path
import subprocess

def extract_audio_fast(video_path, output_dir=None):
    """从视频中快速提取音频（直接复制音频流）"""
    video_file = Path(video_path)

    if not video_file.exists():
        print(f"❌ 文件不存在：{video_path}")
        return False

    # 确定输出目录
    if output_dir:
        output_path = Path(output_dir)
        output_path.mkdir(parents=True, exist_ok=True)
    else:
        output_path = video_file.parent

    # 输出文件名：原文件名.m4a
    output_file = output_path / f"{video_file.stem}.m4a"

    print(f"📹 输入文件：{video_file.name}")
    print(f"🎵 输出文件：{output_file}")
    print(f"⏳ 正在提取音频...")

    # 使用 ffmpeg 直接复制音频流（不重新编码）
    cmd = [
        'ffmpeg',
        '-i', str(video_file),
        '-vn',  # 不要视频
        '-c:a', 'copy',  # 直接复制音频流，不重新编码
        '-y',  # 覆盖已存在文件
        str(output_file)
    ]

    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            check=True
        )
        print(f"✅ 提取成功：{output_file}")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ 提取失败：{e.stderr}")
        return False
    except FileNotFoundError:
        print("❌ 未找到 ffmpeg，请先安装：brew install ffmpeg")
        return False

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法：python3 video_to_audio_fast.py <视频文件> [输出目录]")
        sys.exit(1)

    video_path = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else None

    success = extract_audio_fast(video_path, output_dir)
    sys.exit(0 if success else 1)
