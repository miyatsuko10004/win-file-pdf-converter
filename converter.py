import os
import sys
import glob
import argparse
import win32com.client
import gc
import shutil
import logging
from pathlib import Path
from datetime import datetime
from dotenv import load_dotenv

# --- COM定数定義 ---
ppSaveAsPDF = 32
xlTypePDF = 0
xlSheetVisible = -1  # Excelの表示シート
wdFormatPDF = 17

def setup_logger(output_dir):
    """
    ロガーの設定：コンソール出力とファイル出力の両方を行う
    """
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    log_file_path = output_dir / f"conversion_log_{timestamp}.txt"

    logger = logging.getLogger("PDFConverter")
    logger.setLevel(logging.INFO)
    
    if logger.hasHandlers():
        logger.handlers.clear()

    # ファイルハンドラ (ログファイルへの書き出し)
    fh = logging.FileHandler(log_file_path, encoding='utf-8')
    fh.setLevel(logging.INFO)
    file_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    fh.setFormatter(file_formatter)

    # ストリームハンドラ (コンソールへの表示)
    sh = logging.StreamHandler()
    sh.setLevel(logging.INFO)
    console_formatter = logging.Formatter('%(message)s')
    sh.setFormatter(console_formatter)

    logger.addHandler(fh)
    logger.addHandler(sh)

    return logger, log_file_path

def move_to_done(file_path, done_folder, logger):
    """
    処理完了ファイルをdoneフォルダへ移動
    """
    try:
        dst_path = done_folder / file_path.name
        if dst_path.exists():
            os.remove(dst_path) # 上書きのため既存削除
        shutil.move(str(file_path), str(dst_path))
    except Exception as e:
        logger.warning(f"  [警告] ファイル移動失敗: {file_path.name} -> {e}")

def convert_ppt_to_pdf(target_folder, output_folder, logger):
    """ PowerPoint変換 """
    files = list(Path(target_folder).glob("*.pptx")) + \
            list(Path(target_folder).glob("*.pptm")) + \
            list(Path(target_folder).glob("*.ppt"))
    
    stats = {'success': 0, 'skip': 0, 'error': 0}
    if not files:
        return stats

    done_folder = target_folder / "done"
    done_folder.mkdir(exist_ok=True)

    logger.info(f"--- PowerPoint変換開始: {len(files)}件 ---")
    
    powerpoint = None
    try:
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    except Exception as e:
        logger.error(f"PowerPoint起動失敗: {e}")
        return stats

    try:
        for i, file_path in enumerate(files):
            if i % 10 == 0:
                logger.info(f"PowerPoint 処理中... {i+1}/{len(files)}")

            abs_path = str(file_path.resolve())
            
            if output_folder:
                pdf_path = str((output_folder / file_path.with_suffix('.pdf').name).resolve())
            else:
                pdf_path = str(file_path.with_suffix('.pdf').resolve())

            if os.path.exists(pdf_path):
                logger.info(f"[スキップ] PDF既存: {file_path.name}")
                stats['skip'] += 1
                continue
            
            deck = None
            success = False
            try:
                deck = powerpoint.Presentations.Open(abs_path, WithWindow=False)
                deck.SaveAs(pdf_path, ppSaveAsPDF)
                logger.info(f"[成功] {file_path.name}")
                success = True
                stats['success'] += 1
            except Exception as e:
                logger.error(f"[エラー] {file_path.name}: {e}")
                stats['error'] += 1
            finally:
                if deck:
                    try:
                        deck.Close()
                    except:
                        pass
                    del deck
                gc.collect()

            if success:
                move_to_done(file_path, done_folder, logger)

    finally:
        if powerpoint:
            try:
                powerpoint.Quit()
            except:
                pass
            del powerpoint
            gc.collect()

    logger.info("--- PowerPoint変換終了 ---\n")
    return stats


def convert_excel_to_pdf(target_folder, output_folder, logger):
    """ Excel変換 (強化版: ダイアログ抑制・非表示シート回避) """
    files = list(Path(target_folder).glob("*.xlsx")) + \
            list(Path(target_folder).glob("*.xlsm")) + \
            list(Path(target_folder).glob("*.xls"))
    
    stats = {'success': 0, 'skip': 0, 'error': 0}
    if not files:
        return stats

    done_folder = target_folder / "done"
    done_folder.mkdir(exist_ok=True)

    logger.info(f"--- Excel変換開始: {len(files)}件 ---")

    excel = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False     # 警告抑制
        excel.AskToUpdateLinks = False  # リンク更新確認抑制
        excel.ScreenUpdating = False    # 描画停止
    except Exception as e:
        logger.error(f"Excel起動失敗: {e}")
        return stats

    try:
        for i, file_path in enumerate(files):
            if i % 10 == 0:
                logger.info(f"Excel 処理中... {i+1}/{len(files)}")

            abs_path = str(file_path.resolve())
            
            if output_folder:
                pdf_path = str((output_folder / file_path.with_suffix('.pdf').name).resolve())
            else:
                pdf_path = str(file_path.with_suffix('.pdf').resolve())

            if os.path.exists(pdf_path):
                logger.info(f"[スキップ] PDF既存: {file_path.name}")
                stats['skip'] += 1
                continue

            wb = None
            success = False
            try:
                # ダイアログを出させない強力なOpen設定
                wb = excel.Workbooks.Open(
                    abs_path, 
                    UpdateLinks=0, 
                    ReadOnly=True, 
                    IgnoreReadOnlyRecommended=True,
                    CorruptLoad=1
                )

                # 表示されているシートのみを抽出
                visible_sheets = []
                for ws in wb.Worksheets:
                    if ws.Visible == xlSheetVisible:
                        # 印刷設定の自動調整
                        if not ws.PageSetup.PrintArea:
                            ws.PageSetup.Zoom = False
                            ws.PageSetup.FitToPagesWide = 1
                            ws.PageSetup.FitToPagesTall = False
                        visible_sheets.append(ws.Name)
                
                if visible_sheets:
                    # 可視シートのみを選択してPDF化
                    wb.Worksheets(visible_sheets).Select()
                    wb.ActiveSheet.ExportAsFixedFormat(xlTypePDF, pdf_path, IgnorePrintAreas=False)
                    
                    logger.info(f"[成功] {file_path.name}")
                    success = True
                    stats['success'] += 1
                else:
                    logger.warning(f"[警告] {file_path.name}: 表示可能なシートがありません")
                    stats['error'] += 1

            except Exception as e:
                err_msg = str(e)
                if "Password" in err_msg:
                    logger.error(f"[パスワード保護] {file_path.name}: 開けませんでした")
                else:
                    logger.error(f"[エラー] {file_path.name}: {e}")
                stats['error'] += 1
            finally:
                if wb:
                    try:
                        wb.Close(SaveChanges=False)
                    except:
                        pass
                    del wb
                gc.collect()
            
            if success:
                move_to_done(file_path, done_folder, logger)

    finally:
        if excel:
            try:
                excel.ScreenUpdating = True
                excel.DisplayAlerts = True
                excel.Quit()
            except:
                pass
            del excel
            gc.collect()

    logger.info("--- Excel変換終了 ---\n")
    return stats


def convert_word_to_pdf(target_folder, output_folder, logger):
    """ Word変換 """
    files = list(Path(target_folder).glob("*.docx")) + \
            list(Path(target_folder).glob("*.docm")) + \
            list(Path(target_folder).glob("*.doc"))

    stats = {'success': 0, 'skip': 0, 'error': 0}
    if not files:
        return stats

    done_folder = target_folder / "done"
    done_folder.mkdir(exist_ok=True)

    logger.info(f"--- Word変換開始: {len(files)}件 ---")

    word = None
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False
    except Exception as e:
        logger.error(f"Word起動失敗: {e}")
        return stats

    try:
        for i, file_path in enumerate(files):
            if i % 10 == 0:
                logger.info(f"Word 処理中... {i+1}/{len(files)}")

            abs_path = str(file_path.resolve())

            if output_folder:
                pdf_path = str((output_folder / file_path.with_suffix('.pdf').name).resolve())
            else:
                pdf_path = str(file_path.with_suffix('.pdf').resolve())

            if os.path.exists(pdf_path):
                logger.info(f"[スキップ] PDF既存: {file_path.name}")
                stats['skip'] += 1
                continue

            doc = None
            success = False
            try:
                doc = word.Documents.Open(abs_path)
                doc.SaveAs2(pdf_path, FileFormat=wdFormatPDF)
                logger.info(f"[成功] {file_path.name}")
                success = True
                stats['success'] += 1
            except Exception as e:
                logger.error(f"[エラー] {file_path.name}: {e}")
                stats['error'] += 1
            finally:
                if doc:
                    try:
                        doc.Close()
                    except:
                        pass
                    del doc
                gc.collect()
            
            if success:
                move_to_done(file_path, done_folder, logger)

    finally:
        if word:
            try:
                word.Quit()
            except:
                pass
            del word
            gc.collect()

    logger.info("--- Word変換終了 ---\n")
    return stats


def main():
    load_dotenv()

    parser = argparse.ArgumentParser(description='指定フォルダ内のPPT/Excel/WordファイルをPDFに一括変換し、完了ファイルをdoneフォルダに移動します。')
    parser.add_argument('folder', type=str, nargs='?', help='変換したいファイルが入っているフォルダのパス')
    parser.add_argument('--output', '-o', type=str, help='PDFの出力先フォルダ', default=None)
    args = parser.parse_args()

    folder_str = args.folder or os.getenv('INPUT_FOLDER')

    if not folder_str:
        print("エラー: 変換対象フォルダが指定されていません。引数か.envで指定してください。")
        sys.exit(1)

    target_path = Path(folder_str)

    if not target_path.exists():
        print(f"エラー: 指定されたフォルダが存在しません -> {target_path}")
        sys.exit(1)

    # 出力先の設定
    output_str = args.output or os.getenv('OUTPUT_FOLDER')
    if output_str:
        output_path = Path(output_str)
        try:
            output_path.mkdir(parents=True, exist_ok=True)
            log_dir = output_path 
        except Exception as e:
             print(f"エラー: 出力フォルダ作成失敗 {e}")
             sys.exit(1)
    else:
        output_path = None
        log_dir = target_path

    # ロガーセットアップ
    logger, log_file = setup_logger(log_dir)

    logger.info(f"=== 処理開始: {datetime.now()} ===")
    logger.info(f"対象フォルダ: {target_path.resolve()}")
    if output_path:
        logger.info(f"PDF出力先: {output_path.resolve()}")
    logger.info(f"ログファイル: {log_file}")
    logger.info("--------------------------------------------------\n")
    
    # --- 実行 ---
    ppt_stats = convert_ppt_to_pdf(target_path, output_path, logger)
    xls_stats = convert_excel_to_pdf(target_path, output_path, logger)
    doc_stats = convert_word_to_pdf(target_path, output_path, logger)
    
    # --- 集計 ---
    total_success = ppt_stats['success'] + xls_stats['success'] + doc_stats['success']
    total_skip = ppt_stats['skip'] + xls_stats['skip'] + doc_stats['skip']
    total_error = ppt_stats['error'] + xls_stats['error'] + doc_stats['error']

    logger.info("==================================================")
    logger.info("                最終処理結果サマリー               ")
    logger.info("==================================================")
    logger.info(f"  成功 (PDF作成・移動): {total_success} 件")
    logger.info(f"  スキップ (PDF既存)  : {total_skip} 件")
    logger.info(f"  エラー              : {total_error} 件")
    logger.info("--------------------------------------------------")
    logger.info(f"  PowerPoint -> 成功: {ppt_stats['success']}, エラー: {ppt_stats['error']}")
    logger.info(f"  Excel      -> 成功: {xls_stats['success']}, エラー: {xls_stats['error']}")
    logger.info(f"  Word       -> 成功: {doc_stats['success']}, エラー: {doc_stats['error']}")
    logger.info("==================================================")
    
    print(f"\nすべての処理が完了しました。ログを確認してください: {log_file}")

if __name__ == "__main__":
    main()
