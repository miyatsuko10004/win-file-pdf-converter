# コーディング規約とスタイル

このプロジェクトでは、以下のコーディング規約とスタイルが採用されています。

## 言語

*   Python

## 全般的な規約

*   **命名規則**:
    *   関数名、変数名にはスネークケース (`snake_case`) を使用します。
    *   クラス名にはキャメルケース (`CamelCase`) を使用します。
*   **コメント**:
    *   コード内のコメントは日本語で記述されています。処理の内容や意図を簡潔に説明します。
*   **パス操作**:
    *   `pathlib.Path` オブジェクトを使用してファイルパスを扱います。これにより、OSに依存しないパス操作と可読性の向上が図られています。
*   **引数と環境変数**:
    *   コマンドライン引数 (`argparse`) と `.env` ファイルからの環境変数 (`python-dotenv`) を使用して設定を行います。コマンドライン引数が環境変数よりも優先されます。
*   **エラーハンドリング**:
    *   `try-except` ブロックを使用して、予期されるエラー（例: Officeアプリケーションの起動失敗、ファイル変換エラー）を捕捉し、ユーザーフレンドリーなメッセージを出力します。

## 特定の規約

*   **Windows COMオートメーション**:
    *   `pywin32` ライブラリの `win32com.client.Dispatch` を使用して、Microsoft Officeアプリケーションを操作します。
    *   Officeアプリケーションの `Visible` プロパティを `False` に設定したり、`DisplayAlerts` を `False` に設定したりして、バックグラウンドでの操作を試みています。
*   **PDF変換**:
    *   PowerPointの場合は `Presentations.Open` と `SaveAs(..., 32)` (ppSaveAsPDF) を使用。
    *   Excelの場合は `Workbooks.Open` を使用。印刷範囲が設定されていないシートに対しては `FitToPagesWide=1` (横幅合わせ) を適用し、その後 `Worksheets.Select()`、`ActiveSheet.ExportAsFixedFormat(0, ..., IgnorePrintAreas=False)` を使用。
    *   Wordの場合は `Documents.Open` と `SaveAs2(..., FileFormat=17)` (wdFormatPDF) を使用。
    *   既に変換後のPDFファイルが存在する場合、変換処理をスキップします。

## フォーマットとリンティング

*   明示的なフォーマッター (例: Black, autopep8) やリンター (例: Flake8, Pylint) の設定は確認できませんでした。ただし、コード全体はPEP 8に準拠するように書かれていると推測されます。新しいコードを追加する際は、既存のコードスタイルに合わせることが推奨されます。
