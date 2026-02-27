"""pptx2jpg.py - PowerPointファイルをJPEGに変換するCLIツール

PowerPointのCOMオブジェクトを使用してスライドをJPEG画像にエクスポートする。
WSL側のAI Agentから呼び出される想定で、WSLパスの自動変換機能を備える。

使い方:
    python pptx2jpg.py -i presentation.pptx -o slide.jpg -p 1
    python pptx2jpg.py -i presentation.pptx -o output_dir -a
"""

import argparse
import os
import subprocess
import sys


def to_windows_path(path):
    """WSLパスをWindowsパスに変換する。既にWindowsパスの場合はそのまま返す。

    Args:
        path: ファイルまたはフォルダのパス

    Returns:
        Windowsパス文字列
    """
    # 既にWindowsパス（ドライブレターまたはUNC）の場合はそのまま
    if len(path) >= 2 and path[1] == ":":
        return path
    if path.startswith("\\\\"):
        return path

    # Unixスタイルのパス（/で始まる）→ wslpathで変換
    if path.startswith("/"):
        try:
            result = subprocess.run(
                ["wslpath", "-w", path],
                capture_output=True,
                text=True,
                check=True,
            )
            return result.stdout.strip()
        except FileNotFoundError:
            # wslpathが無い＝Windows上で直接実行されている可能性
            # そのまま返して後続の処理に委ねる
            return path
        except subprocess.CalledProcessError as e:
            print(f"エラー: WSLパスの変換に失敗しました: {path}", file=sys.stderr)
            print(f"  wslpath stderr: {e.stderr.strip()}", file=sys.stderr)
            sys.exit(1)

    # 相対パスの場合は絶対パスに変換
    return os.path.abspath(path)


def export_single_slide(input_path, output_path, page):
    """指定スライドをJPEGにエクスポートする。

    Args:
        input_path: PowerPointファイルのWindowsパス
        output_path: 出力JPEGファイルのWindowsパス
        page: スライド番号（1始まり）
    """
    import win32com.client

    ppt = None
    presentation = None

    try:
        ppt = win32com.client.Dispatch("PowerPoint.Application")
        # ウィンドウ非表示で開く (FileName, ReadOnly, HasTitle, WithWindow)
        presentation = ppt.Presentations.Open(input_path, True, False, False)

        slide_count = presentation.Slides.Count
        if page < 1 or page > slide_count:
            print(
                f"エラー: スライド番号 {page} は範囲外です（1〜{slide_count}）",
                file=sys.stderr,
            )
            sys.exit(1)

        # 出力先ディレクトリが無ければ作成
        output_dir = os.path.dirname(output_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir, exist_ok=True)

        slide = presentation.Slides(page)
        slide.Export(output_path, "JPG")

        print(f"Exported: {output_path}")

    except Exception as e:
        print(f"エラー: {e}", file=sys.stderr)
        sys.exit(1)
    finally:
        if presentation is not None:
            try:
                presentation.Close()
            except Exception:
                pass
        if ppt is not None:
            try:
                ppt.Quit()
            except Exception:
                pass


def export_all_slides(input_path, output_dir):
    """全スライドをJPEGにエクスポートする。

    Args:
        input_path: PowerPointファイルのWindowsパス
        output_dir: 出力フォルダのWindowsパス
    """
    import win32com.client

    ppt = None
    presentation = None

    try:
        # 出力フォルダが無ければ作成
        if not os.path.exists(output_dir):
            os.makedirs(output_dir, exist_ok=True)

        ppt = win32com.client.Dispatch("PowerPoint.Application")
        # ウィンドウ非表示で開く (FileName, ReadOnly, HasTitle, WithWindow)
        presentation = ppt.Presentations.Open(input_path, True, False, False)

        slide_count = presentation.Slides.Count
        presentation.Export(output_dir, "JPG")

        print(f"Exported {slide_count} slides to: {output_dir}")

    except Exception as e:
        print(f"エラー: {e}", file=sys.stderr)
        sys.exit(1)
    finally:
        if presentation is not None:
            try:
                presentation.Close()
            except Exception:
                pass
        if ppt is not None:
            try:
                ppt.Quit()
            except Exception:
                pass


def parse_args(argv=None):
    """コマンドライン引数をパースする。"""
    parser = argparse.ArgumentParser(
        description="PowerPointファイルをJPEGに変換する"
    )
    parser.add_argument(
        "-i",
        "--input",
        required=True,
        help="変換するPowerPointファイル",
    )
    parser.add_argument(
        "-o",
        "--output",
        required=True,
        help="-pの時はJPEGファイル名、-aの時は出力フォルダ名",
    )

    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument(
        "-p",
        "--page",
        type=int,
        help="変換するスライド番号（1始まり）",
    )
    mode.add_argument(
        "-a",
        "--all",
        action="store_true",
        help="すべてのスライドをJPEGに変換",
    )

    return parser.parse_args(argv)


def main():
    args = parse_args()

    # パスをWindowsパスに変換（WSLから呼ばれた場合に対応）
    input_path = to_windows_path(args.input)
    output_path = to_windows_path(args.output)

    # 入力ファイルの存在チェック
    if not os.path.exists(input_path):
        print(f"エラー: 入力ファイルが見つかりません: {input_path}", file=sys.stderr)
        sys.exit(1)

    if args.page is not None:
        export_single_slide(input_path, output_path, args.page)
    else:
        export_all_slides(input_path, output_path)


if __name__ == "__main__":
    main()
