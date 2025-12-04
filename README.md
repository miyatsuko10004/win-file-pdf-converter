# Office to PDF Converter

指定したフォルダ内にある PowerPoint (.pptx, .ppt) および Excel (.xlsx, .xls) ファイルを自動的に PDF に変換するツール。
Windows にインストールされている Microsoft Office をバックグラウンドで操作して変換を行うため、レイアウト崩れを最小限抑える。

## 前提条件

* **OS**: Windows 10 / 11
* **Software**: Microsoft Office (PowerPoint, Excel) がインストールされていること
* **Tool**: [uv](https://github.com/astral-sh/uv)

## 使い方

変換したいファイルが入っているフォルダを用意し、以下のコマンドを実行する。
`uv` が自動的に必要なライブラリ (`pywin32`) を用意して実行する。

```bash
uv run --with pywin32 converter.py "C:\Path\To\Your\TargetFolder" --output "C:\Path\To\OutputFolder"
```

※ 出力先フォルダを指定しない場合は、元のファイルと同じフォルダに保存される。

※ フォルダパスにスペースが含まれる場合は、ダブルクォーテーション " で囲う。

## 注意事項
* Excelの変換範囲: Excelファイルは、各ファイル内で設定されている「印刷範囲」または「改ページプレビュー」の設定に基づいてPDF化されます。意図しない列のはみ出しを防ぐため、事前にExcel側で印刷範囲を確認することを推奨。
* 実行中の操作: スクリプト実行中に、バックグラウンドでOfficeアプリが開閉を繰り返す。誤作動を防ぐため、実行中はExcelやPowerPointの手動操作を控えることを推奨。
* エラー処理: パスワード付きのファイルや破損したファイルが含まれている場合、そのファイルはスキップ（エラー表示）され、処理は継続。
