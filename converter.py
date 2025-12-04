import os
import sys
import glob
import argparse
import win32com.client
from pathlib import Path

def convert_ppt_to_pdf(target_folder):
    """
    指定フォルダ内のPowerPointファイルをPDFに変換します。
    """
    # 検索パターン（.pptx と .ppt）
    files = list(Path(target_folder).glob("*.pptx")) + list(Path(target_folder).glob("*.ppt"))
    
    if not files:
        print("-> PowerPointファイルは見つかりませんでした。")
        return

    print(f"--- PowerPoint変換開始: {len(files)}件 ---")
    
    try:
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        # 処理高速化・干渉防止のためウィンドウを最小化（完全非表示はバージョンにより不安定なため）
        # powerpoint.Visible = True 
    except Exception as e:
        print(f"PowerPointの起動に失敗しました: {e}")
        return

    for file_path in files:
        abs_path = str(file_path.resolve())
        pdf_path = str(file_path.with_suffix('.pdf').resolve())

        try:
            # 既にPDFが存在する場合はスキップ（上書きしたい場合はコメントアウト）
            if os.path.exists(pdf_path):
                print(f"[スキップ] 既に存在します: {file_path.name}")
                continue

            deck = powerpoint.Presentations.Open(abs_path, WithWindow=False)
            
            # formatType 32 = ppSaveAsPDF
            deck.SaveAs(pdf_path, 32)
            deck.Close()
            print(f"[成功] {file_path.name}")
        except Exception as e:
            print(f"[エラー] {file_path.name}: {e}")

    powerpoint.Quit()
    print("--- PowerPoint変換終了 ---\n")


def convert_excel_to_pdf(target_folder):
    """
    指定フォルダ内のExcelファイルをPDFに変換します。
    ※印刷範囲設定を反映し、全シートを1つのPDFに出力します。
    """
    files = list(Path(target_folder).glob("*.xlsx")) + list(Path(target_folder).glob("*.xls"))
    
    if not files:
        print("-> Excelファイルは見つかりませんでした。")
        return

    print(f"--- Excel変換開始: {len(files)}件 ---")

    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
    except Exception as e:
        print(f"Excelの起動に失敗しました: {e}")
        return

    for file_path in files:
        abs_path = str(file_path.resolve())
        pdf_path = str(file_path.with_suffix('.pdf').resolve())

        try:
            if os.path.exists(pdf_path):
                print(f"[スキップ] 既に存在します: {file_path.name}")
                continue

            wb = excel.Workbooks.Open(abs_path)
            
            # 【重要】全シートをPDF対象にするため、すべてのシートを選択状態にする
            # これを行わないと、保存時に開いていたシートしかPDFにならない場合があります
            wb.Worksheets.Select()
            
            # Type=0 (xlTypePDF), IgnorePrintAreas=False (デフォルト)
            # IgnorePrintAreas=False なので、Excel側で設定した「印刷範囲」が守られます
            wb.ActiveSheet.ExportAsFixedFormat(0, pdf_path, IgnorePrintAreas=False)
            
            wb.Close(False)
            print(f"[成功] {file_path.name}")
        except Exception as e:
            print(f"[エラー] {file_path.name}: {e}")

    excel.Quit()
    print("--- Excel変換終了 ---\n")


def main():
    parser = argparse.ArgumentParser(description='指定フォルダ内のPPT/ExcelファイルをPDFに一括変換します。')
    parser.add_argument('folder', type=str, help='変換したいファイルが入っているフォルダのパス')
    args = parser.parse_args()

    target_path = Path(args.folder)

    if not target_path.exists():
        print(f"エラー: 指定されたフォルダが存在しません -> {target_path}")
        sys.exit(1)

    print(f"対象フォルダ: {target_path.resolve()}\n")
    
    convert_ppt_to_pdf(target_path)
    convert_excel_to_pdf(target_path)
    
    print("すべての処理が完了しました。")

if __name__ == "__main__":
    main()
