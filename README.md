# pptx2jpg

PowerPointファイルをJPEG画像に変換するCLIツール。

AI AgentがPowerPointファイルの見た目（文字の重なり、罫線の不表示等）を画像で確認できるようにするために作成。PowerPoint本体のCOMオブジェクトを使用して高品質なJPEGエクスポートを行う。

## 前提条件

- **WSL2**（推奨）またはWindows環境
- **Microsoft PowerPoint** がインストール済み
- **WinPython** (`C:\tools\WPy64-31380\python\python.exe`)
  - `pywin32`（`win32com.client`）が必要（WinPythonに同梱済み）

## インストール

WSL2上で以下を実行する:

```bash
make install
```

以下の2つがインストールされる:

| ファイル | インストール先 | 説明 |
| --- | --- | --- |
| `pptx2jpg.py` | `/mnt/c/tools/` | Windows側で動作する本体スクリプト |
| `pptx2jpg` | `~/.local/bin/` | WSL2から呼び出すラッパーシェルスクリプト |

インストール先を変更する場合:

```bash
make install TARGET=/mnt/c/mytools BINDIR=~/bin
```

アンインストール:

```bash
make clean
```

## 使い方

### WSL2から使う（推奨）

`~/.local/bin` にPATHが通っていれば、`pptx2jpg` コマンドとして直接実行できる。

特定のスライドを変換:

```bash
pptx2jpg -i presentation.pptx -o slide1.jpg -p 1
```

すべてのスライドを変換:

```bash
pptx2jpg -i presentation.pptx -o output_dir -a
```

`-i` / `-o` に渡すパスはWSLパス（`/home/...` や `/mnt/c/...`）でOK（自動的にWindowsパスに変換される）。

### Windows（PowerShell）から使う

```powershell
C:\tools\WPy64-31380\python\python.exe C:\tools\pptx2jpg.py -i presentation.pptx -o slide1.jpg -p 1
```

```powershell
C:\tools\WPy64-31380\python\python.exe C:\tools\pptx2jpg.py -i presentation.pptx -o output_dir -a
```

全スライドがPowerPoint標準の命名規則（`スライド1.JPG`, `スライド2.JPG`, ...）で出力される。

## オプション

| オプション | 説明 |
| --- | --- |
| `-i / --input` | 変換するPowerPointファイル（必須） |
| `-o / --output` | `-p` の時はJPEGファイル名、`-a` の時は出力フォルダ名（必須） |
| `-p / --page` | 変換するスライド番号（1始まり） |
| `-a / --all` | すべてのスライドをJPEGに変換 |

`-p` と `-a` はどちらか一方を必ず指定する（同時指定は不可）。

## 環境変数

ラッパースクリプト `pptx2jpg` は以下の環境変数でカスタマイズできる:

| 変数 | デフォルト | 説明 |
| --- | --- | --- |
| `PYTHON_EXE` | `/mnt/c/tools/WPy64-31380/python/python.exe` | Windows側のPython実行ファイル |
| `PPTX2JPG_PY` | `/mnt/c/tools/pptx2jpg.py` | pptx2jpg.py のパス |
