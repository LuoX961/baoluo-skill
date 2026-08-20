#!/bin/bash
# baoluo-pdf-scan-to-markdown 依赖安装：创建 OCR 虚拟环境并安装 pyobjc
# 依赖 macOS 12+（Vision 框架）与系统 Python 3
set -eu

VENV="${HOME}/.venvs/ocr"
PYTHON_BIN="$(command -v python3)"

echo "→ 使用 Python: ${PYTHON_BIN}"
if [ ! -d "$VENV" ]; then
  echo "→ 创建虚拟环境: ${VENV}"
  "$PYTHON_BIN" -m venv "$VENV"
fi

echo "→ 安装/升级 pip"
"$VENV/bin/python" -m pip install --upgrade pip -q

echo "→ 安装 pyobjc（Vision OCR + Quartz PDF 渲染）"
"$VENV/bin/python" -m pip install pyobjc-framework-Vision pyobjc-framework-Quartz -q

echo "→ 验证"
"$VENV/bin/python" - <<'EOF'
import Vision, Quartz
print("✓ Vision OCR 接口可用（macOS 原生识别引擎）")
EOF

echo "完成。后续请使用: ${VENV}/bin/python"
