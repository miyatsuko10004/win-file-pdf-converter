import os
import sys
import glob
import argparse
import win32com.client
import gc  # ガベージコレクション用に追加
from pathlib import Path
from dotenv import load_dotenv

# COM定数の定義
ppSaveAsPDF = 32
xlTypePDF = 0
wdFormatPDF = 17

def convert_ppt_to_pdf(target_folder, output_folder=None):
    """
    指定フォルダ内のPowerPointファイルをPDFに変換します。
    """
    files = list(Path(target_folder).glob("*.pptx")) + \
            list(Path(target_folder).glob("*.pptm")) + \
            list(Path(target_folder).glob("*.ppt"))
    
    if not files:
        print("-> PowerPointファイルは見つかりませんでした。")
        return

    print(f"--- PowerPoint変換開始: {len(files)}件 ---")
    
    powerpoint = None
    try:
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        # powerpoint.Visible = True 
    except Exception as e:
        print(f"PowerPointの起動に失敗しました: {e}")
        return

    try:
        for i, file_path in enumerate(files):
            # 進捗表示（1000件あると進んでるか不安になるため）
            if i % 10 == 0:
                print(f"処理中... {i+1}/{len(files)}")

            abs_path = str(file_path.resolve())
            
            if output_folder:
                pdf_path = str((output_folder / file_path.with_suffix('.pdf').name).resolve())
            else:
                pdf_path = str(file_path.with_suffix('.pdf').resolve())

            if os.path.exists(pdf_path):
                print(f"[スキップ] 既に存在します: {file_path.name}")
                continue
            
            deck = None
            try:
                deck = powerpoint.Presentations.Open(abs_path, WithWindow=False)
                deck.SaveAs(pdf_path, ppSaveAsPDF)
                print(f"[成功] {file_path.name}")
            except Exception as e:
                print(f"[エラー] {file_path.name}: {e}")
            finally:
                # 確実に閉じてメモリ解放
                if deck:
                    try:
                        deck.Close()
                    except:
                        pass
                    del deck  # 参照を削除
                
                # ガベージコレクションを強制実行してメモリリークを防ぐ
                gc.collect()

    finally:
        # ループ中にエラーが起きても必ず終了させる
        if powerpoint:
            try:
                powerpoint.Quit()
            except:
                pass
            del powerpoint
            gc.collect()

    print("--- PowerPoint変換終了 ---\n")


def convert_excel_to_pdf(target_folder, output_folder=None):
    """
    指定フォルダ内のExcelファイルをPDFに変換します。
    """
    files = list(Path(target_folder).glob("*.xlsx")) + \
            list(Path(target_folder).glob("*.xlsm")) + \
            list(Path(target_folder).glob("*.xls"))
    
    if not files:
        print("-> Excelファイルは見つかりませんでした。")
        return

    print(f"--- Excel変換開始: {len(files)}件 ---")

    excel = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
    except Exception as e:
        print(f"Excelの起動に失敗しました: {e}")
        return

    try:
        for i, file_path in enumerate(files):
            if i % 10 == 0:
                print(f"処理中... {i+1}/{len(files)}")

            abs_path = str(file_path.resolve())
            
            if output_folder:
                pdf_path = str((output_folder / file_path.with_suffix('.pdf').name).resolve())
            else:
                pdf_path = str(file_path.with_suffix('.pdf').resolve())

            if os.path.exists(pdf_path):
                print(f"[スキップ] 既に存在します: {file_path.name}")
                continue

            wb = None
            try:
                wb = excel.Workbooks.Open(abs_path)

                for ws in wb.Worksheets:
                    if not ws.PageSetup.PrintArea:
                        ws.PageSetup.Zoom = False
                        ws.PageSetup.FitToPagesWide = 1
                        ws.PageSetup.FitToPagesTall = False
                
                wb.Worksheets.Select()
                wb.ActiveSheet.ExportAsFixedFormat(xlTypePDF, pdf_path, IgnorePrintAreas=False)
                print(f"[成功] {file_path.name}")
            except Exception as e:
                print(f"[エラー] {file_path.name}: {e}")
            finally:
                if wb:
                    try:
                        wb.Close(False)
                    except:
                        pass
                    del wb
                gc.collect()

    finally:
        if excel:
            try:
                excel.Quit()
            except:
                pass
            del excel
            gc.collect()

    print("--- Excel変換終了 ---\n")


def convert_word_to_pdf(target_folder, output_folder=None):
    """
    指定フォルダ内のWordファイルをPDFに変換します。
    """
    files = list(Path(target_folder).glob("*.docx")) + \
            list(Path(target_folder).glob("*.docm")) + \
            list(Path(target_folder).glob("*.doc"))

    if not files:
        print("-> Wordファイルは見つかりませんでした。")
        return

    print(f"--- Word変換開始: {len(files)}件 ---")

    word = None
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False
    except Exception as e:
        print(f"Wordの起動に失敗しました: {e}")
        return

    try:
        for i, file_path in enumerate(files):
            if i % 10 == 0:
                print(f"処理中... {i+1}/{len(files)}")

            abs_path = str(file_path.resolve())

            if output_folder:
                pdf_path = str((output_folder / file_path.with_suffix('.pdf').name).resolve())
            else:
                pdf_path = str(file_path.with_suffix('.pdf').resolve())

            if os.path.exists(pdf_path):
                print(f"[スキップ] 既に存在します: {file_path.name}")
                continue

            doc = None
            try:
                doc = word.Documents.Open(abs_path)
                doc.SaveAs2(pdf_path, FileFormat=wdFormatPDF)
                print(f"[成功] {file_path.name}")
            except Exception as e:
                print(f"[エラー] {file_path.name}: {e}")
            finally:
                if doc:
                    try:
                        doc.Close()
                    except:
                        pass
                    del doc
                gc.collect()
    finally:
        if word:
            try:
                word.Quit()
            except:
                pass
            del word
            gc.collect()

    print("--- Word変換終了 ---\n")


def main():
    load_dotenv()

    parser = argparse.ArgumentParser(description='指定フォルダ内のPPT/Excel/WordファイルをPDFに一括変換します。')
    parser.add_argument('folder', type=str, nargs='?', help='変換したいファイルが入っているフォルダのパス')
    parser.add_argument('--output', '-o', type=str, help='PDFの出力先フォルダ', default=None)
    args = parser.parse_args()

    folder_str = args.folder or os.getenv('INPUT_FOLDER')

    if not folder_str:
        print("エラー: 変換対象フォルダが指定されていません。")
        sys.exit(1)

    target_path = Path(folder_str)

    if not target_path.exists():
        print(f"エラー: 指定されたフォルダが存在しません -> {target_path}")
        sys.exit(1)

    output_str = args.output or os.getenv('OUTPUT_FOLDER')
    output_path = None
    
    if output_str:
        output_path = Path(output_str)
        if not output_path.exists():
            try:
                output_path.mkdir(parents=True, exist_ok=True)
                print(f"出力フォルダを作成しました: {output_path.resolve()}")
            except Exception as e:
                print(f"エラー: 出力フォルダの作成に失敗しました -> {e}")
                sys.exit(1)
        print(f"出力先フォルダ: {output_path.resolve()}")

    print(f"対象フォルダ: {target_path.resolve()}\n")
    
    convert_ppt_to_pdf(target_path, output_path)
    convert_excel_to_pdf(target_path, output_path)
    convert_word_to_pdf(target_path, output_path)
    
    print("すべての処理が完了しました。")

if __name__ == "__main__":
    main()
