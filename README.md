# pptx2jpg

PowerPointファイルをJPEG画像に変換するCLIツール。

AI AgentがPowerPointファイルの見た目（文字の重なり、罫線の不表示等）を画像で確認できるようにするために作成。PowerPoint本体のCOMオブジェクトを使用して高品質なJPEGエクスポートを行う。

## 前提条件

- **Windows環境**（PowerPointのCOM操作のため）
- **Microsoft PowerPoint** がインストール済み
- **WinPython** (`C:\tools\WPy64-31380\python\python.exe`)
  - `pywin32`（`win32com.client`）が必要（WinPythonに同梱済み）

## インストール

```
install.bat
```

デフォルトで `C:\tools` にインストールされる。フォルダを変更する場合:

```
install.bat D:\mytools
```

アンインストール:

```
install.bat -clean
install.bat -clean D:\mytools
```

## 使い方

### 特定のスライドを変換

```
python pptx2jpg.py -i <PowerPointファイル> -o <出力JPEGファイル> -p <スライド番号>
```

例:

```powershell
C:\tools\WPy64-31380\python\python.exe C:\tools\pptx2jpg.py -i presentation.pptx -o slide1.jpg -p 1
```

### すべてのスライドを変換

```
python pptx2jpg.py -i <PowerPointファイル> -o <出力フォルダ> -a
```

例:

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

## WSLからの利用

WSL側から呼び出す場合、`-i` / `-o` に渡すパスはWSLパス（`/home/...` や `/mnt/c/...`）でもOK（自動的にWindowsパスに変換される）。

ただし、**スクリプト自体のパスはWindowsパス（`C:/tools/pptx2jpg.py`）で指定する**こと。スクリプトパスはPython本体が解釈するため、WSLパス変換の対象外となる。Python実行ファイルはbashが見つけられるよう `/mnt/c/...` 形式で指定する。

```bash
/mnt/c/tools/WPy64-31380/python/python.exe C:/tools/pptx2jpg.py \
  -i /home/user/docs/presentation.pptx \
  -o /home/user/docs/slide1.jpg \
  -p 1
```
