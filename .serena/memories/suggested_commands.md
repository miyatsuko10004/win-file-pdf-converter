# 開発に役立つコマンド

このプロジェクトでの開発および運用に役立つ主要なコマンドを以下に示します。

## パッケージ管理 (uv)

*   **依存関係のインストール**:
    ```bash
    uv sync
    ```
    `pyproject.toml` に基づいて依存関係をインストールします。

*   **開発用依存関係のインストール**:
    ```bash
    uv sync --with dev
    ```
    `pyproject.toml` の `[dependency-groups.dev]` に定義されている開発用依存関係をインストールします。

## テスト

*   **全テストの実行**:
    ```bash
    uv run python -m unittest discover tests
    ```
    `tests` ディレクトリ内のすべてのテストファイルを検出して実行します。

*   **特定のテストファイルの実行**:
    ```bash
    uv run python -m unittest tests/test_converter.py
    ```
    `tests/test_converter.py` のテストを実行します。

## プロジェクトのエントリポイントの実行

*   **変換ツールの実行**:
    ```bash
    uv run --with pywin32 converter.py "C:\Path\To\Your\TargetFolder" --output "C:\Path\To\OutputFolder"
    ```
    指定されたフォルダ内のOfficeファイルをPDFに変換します。`pywin32` はWindowsでのみ必要となるため `--with pywin32` オプションを付けます。
    フォルダパスにスペースが含まれる場合は、ダブルクォーテーションで囲んでください。
    出力先フォルダを省略した場合、入力フォルダと同じ場所にPDFが生成されます。

*   **環境変数を使用した変換ツールの実行**:
    ```bash
    uv run --with pywin32 converter.py
    ```
    `.env` ファイルに `INPUT_FOLDER` および `OUTPUT_FOLDER` が設定されている場合、引数なしで実行できます。

## その他

*   **Python環境のセットアップ**:
    ```bash
    uv python install
    ```
    `uv` を使用してPython環境をセットアップします。

*   **仮想環境の作成**:
    ```bash
    uv venv
    ```
    `.venv` ディレクトリに仮想環境を作成します。

*   **仮想環境の有効化**:
    ```bash
    source .venv/bin/activate
    ```
    （Linux/macOSの場合）

    ```bash
    .venv\Scripts\activate
    ```
    （Windowsの場合）

## Gitコマンド

基本的なGit操作（`git status`, `git add`, `git commit`, `git push`, `git pull` など）は通常のLinux環境と同様に使用できます。

## ファイル操作コマンド

基本的なファイル操作コマンド（`ls`, `cd`, `cp`, `mv`, `rm`, `mkdir`, `grep`, `find` など）は通常のLinux環境と同様に使用できます。
