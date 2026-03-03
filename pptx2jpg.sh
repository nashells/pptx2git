#!/usr/bin/env bash
# pptx2jpg.sh - WSL2からpptx2jpg.pyをWindows Python経由で実行するラッパー
#
# 使い方:
#   pptx2jpg.sh -i presentation.pptx -o slide.jpg -p 1
#   pptx2jpg.sh -i presentation.pptx -o output_dir -a
set -euo pipefail

PYTHON_EXE="${PYTHON_EXE:-/mnt/c/tools/WPy64-31380/python/python.exe}"
PPTX2JPG="${PPTX2JPG_PY:-/mnt/c/tools/pptx2jpg.py}"

# python.exe の存在チェック
if [[ ! -f "$PYTHON_EXE" ]]; then
    echo "エラー: $PYTHON_EXE が見つかりません" >&2
    exit 1
fi

# pptx2jpg.py の存在チェック
if [[ ! -f "$PPTX2JPG" ]]; then
    echo "エラー: $PPTX2JPG が見つかりません" >&2
    echo "  make install で先にインストールしてください" >&2
    exit 1
fi

# Windows Python にパスを渡すため pptx2jpg.py のパスを変換
WIN_SCRIPT="$(wslpath -w "$PPTX2JPG")"

exec "$PYTHON_EXE" "$WIN_SCRIPT" "$@"
