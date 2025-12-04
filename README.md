# Office to PDF Converter

指定したフォルダ内にある PowerPoint (.pptx, .ppt) および Excel (.xlsx, .xls) ファイルを自動的に PDF に変換するツールです。
Windows にインストールされている Microsoft Office をバックグラウンドで操作して変換を行うため、レイアウト崩れを最小限に抑えられます。

## 前提条件 (Requirements)

* **OS**: Windows 10 / 11
* **Software**: Microsoft Office (PowerPoint, Excel) がインストールされていること
* **Python**: Python 3.x

## インストール (Installation)

必要なライブラリ `pywin32` をインストールしてください。

```bash
pip install pywin32
```

使い方 (Usage)
1. 変換したいファイルが入っているフォルダを用意します。
2. 以下のコマンドを実行します。
```
python converter.py "C:\Path\To\Your\TargetFolder"
```

※ フォルダパスにスペースが含まれる場合は、ダブルクォーテーション " で囲ってください。

注意事項 (Notes)
• Excelの変換範囲: Excelファイルは、各ファイル内で設定されている「印刷範囲」または「改ページプレビュー」の設定に基づいてPDF化されます。意図しない列のはみ出しを防ぐため、事前にExcel側で印刷範囲を確認することを推奨します。
• 実行中の操作: スクリプト実行中は、バックグラウンドでOfficeアプリが開閉を繰り返します。誤作動を防ぐため、実行中はExcelやPowerPointの手動操作を控えてください。
• エラー処理: パスワード付きのファイルや破損したファイルが含まれている場合、そのファイルはスキップ（エラー表示）され、処理は継続します。




