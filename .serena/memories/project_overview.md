# プロジェクト概要

このプロジェクトは「Office to PDF Converter」という名称で、Windows環境においてPowerPoint (.pptx, .ppt)、Excel (.xlsx, .xls)、Word (.docx, .docm, .doc) ファイルを自動的にPDFに変換するツールです。変換にはMicrosoft OfficeがインストールされているWindows OSを利用し、COM (Component Object Model) オートメーションを通じてOfficeアプリケーションをバックグラウンドで操作することで、高い精度でのPDF変換を実現しています。

## 目的

*   PowerPoint、Excel、WordファイルをPDFに自動変換し、手動変換の手間を省く。
*   Microsoft Officeの機能を活用することで、レイアウト崩れを最小限に抑えた高品質なPDFを生成する。
*   コマンドライン引数または環境変数を通じて、変換元フォルダと出力先フォルダを柔軟に指定できるようにする。

## 技術スタック

*   **言語**: Python (3.10以上)
*   **依存ライブラリ**:
    *   `pywin32`: Windows COMオートメーションを通じてMicrosoft Officeアプリケーションを操作するために使用。Windows OSでのみ必要。
    *   `python-dotenv`: `.env` ファイルから環境変数をロードするために使用。
    *   `pathlib`: パス操作を容易にする。
    *   `argparse`: コマンドライン引数を解析する。
*   **パッケージマネージャー**: `uv`
*   **ビルドシステム**: `hatchling`
*   **テストフレームワーク**: `unittest`
*   **CI/CD**: GitHub Actions

## コードベースの主要な構成

*   `converter.py`: メインの変換ロジックを実装。PowerPoint、Excel、Wordそれぞれに対する変換関数と、それらを呼び出す`main`関数を含む。
*   `tests/test_converter.py`: `converter.py`のユニットテスト。`win32com`のモック化により、Windows以外の環境でもテスト実行が可能。
*   `pyproject.toml`: プロジェクトのメタデータ、依存関係、ビルド設定を定義。
*   `README.md`: プロジェクトの概要、前提条件、使用方法、注意事項を記述。
*   `.github/workflows/test.yml`: GitHub ActionsによるCI設定。`uv`を用いた依存関係のインストールと`unittest`によるテスト実行を定義。

## 実行環境の前提

*   OS: Windows 10 / 11
*   ソフトウェア: Microsoft Office (PowerPoint, Excel, Word)
*   ツール: `uv` (Pythonパッケージマネージャー)

本プロジェクトは、Microsoft OfficeがインストールされたWindows環境でのみ完全に機能します。テストは `win32com` をモックすることでLinux環境でも実行可能です。
