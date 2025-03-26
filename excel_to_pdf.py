import os
import traceback
import win32com.client

# グローバル変数として、元フォルダと出力フォルダのパスを定義
SOURCE_FOLDER = r"input"
OUTPUT_FOLDER = r"output"

def convert_excel_to_pdf(excel_path, output_folder):
    """
    指定したExcelファイル内の全シートをPDFに変換して保存する関数。
    PDFのファイル名は "{Excelファイルのファイル名}_{シート名}.pdf" 形式。
    """
    # 絶対パスに変換（Excelは絶対パスでの指定が安定のため）
    abs_excel_path = os.path.abspath(excel_path)
    
    # Excelファイルの存在確認
    if not os.path.exists(abs_excel_path):
        print(f"ファイルが存在しません: {abs_excel_path}")
        return

    try:
        # Excelアプリケーションのインスタンス生成
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False       # Excelウィンドウを非表示
        excel.DisplayAlerts = False # ダイアログの自動抑制
    except Exception as e:
        print(f"Excelアプリケーションの作成エラー: {e}")
        return

    workbook = None
    try:
        # Excelファイルをオープン
        workbook = excel.Workbooks.Open(abs_excel_path)
        if workbook is None:
            print(f"Workbookがオープンできませんでした: {abs_excel_path}")
            return

        # 拡張子を除いたExcelファイル名を取得
        base_filename = os.path.splitext(os.path.basename(excel_path))[0]
        
        # ワークブック内の各シートに対して処理
        for sheet in workbook.Worksheets:
            try:
                sheet_name = sheet.Name
                # Windowsのファイル名で使えない文字を除去
                invalid_chars = '<>:"/\\|?*'
                for char in invalid_chars:
                    sheet_name = sheet_name.replace(char, '')
                
                # 出力先のファイルパス生成
                output_file = os.path.join(output_folder, f"{base_filename}_{sheet_name}.pdf")
                
                # シートをPDFにエクスポート（0はPDF形式）
                sheet.ExportAsFixedFormat(0, output_file)
                print(f"変換成功: {output_file}")
            except Exception as e:
                print(f"シート変換エラー (シート名: {sheet.Name}, ファイル: {excel_path}): {e}")
                traceback.print_exc()
                continue  # 次のシートへ
    except Exception as e:
        print(f"ワークブックオープンエラー ({excel_path}): {e}")
        traceback.print_exc()
    finally:
        # ワークブックが正常にオープンしていれば閉じる
        try:
            if workbook:
                workbook.Close(False)
        except Exception as e:
            print(f"ワークブックのクローズ中にエラー: {e}")
        excel.Quit()

def main():
    # 絶対パスに変換（保存時挙動の安定性向上のため）
    abs_output_folder = os.path.abspath(OUTPUT_FOLDER)
    if not os.path.exists(abs_output_folder):
        try:
            os.makedirs(abs_output_folder)
        except Exception as e:
            print(f"出力フォルダ作成エラー ({abs_output_folder}): {e}")
            return

    # 指定フォルダ内（サブディレクトリ含む）の全Excelファイルを処理
    for root, dirs, files in os.walk(SOURCE_FOLDER):
        for file in files:
            if file.lower().endswith(".xlsx"):
                excel_file_path = os.path.join(root, file)
                print(f"処理開始: {excel_file_path}")
                try:
                    convert_excel_to_pdf(excel_file_path, abs_output_folder)
                except Exception as e:
                    print(f"ファイル処理中エラー ({excel_file_path}): {e}")
                    traceback.print_exc()
                    continue

if __name__ == "__main__":
    main()