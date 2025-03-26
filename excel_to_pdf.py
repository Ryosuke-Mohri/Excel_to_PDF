import os
import traceback
import win32com.client

# グローバル変数として、元フォルダと出力フォルダのパスを定義
SOURCE_FOLDER = "input"
OUTPUT_FOLDER = "output"

def convert_excel_to_pdf(excel_path, output_folder):
    
    # 指定したExcelファイル内の全シートをPDFに変換して保存する関数

    # Excelアプリケーションのインスタンス生成
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False       # Excelウィンドウを非表示にする
        excel.DisplayAlerts = False # ダイアログの自動抑制
    except Exception as e:
        print(f"Excelアプリケーションの作成エラー: {e}")
        return

    # Excelファイルをオープン
    try:
        workbook = excel.Workbooks.Open(excel_path)
    except Exception as e:
        print(f"ワークブックオープンエラー ({excel_path}): {e}")
        excel.Quit()
        return

    # Excelファイル名（拡張子を除く）を取得
    base_filename = os.path.splitext(os.path.basename(excel_path))[0]
    
    try:
        # ワークブック内の各シートを処理
        for sheet in workbook.Worksheets:
            try:
                sheet_name = sheet.Name
                # Windowsでファイル名に使えない文字を除去
                invalid_chars = '<>:"/\\|?*'
                for char in invalid_chars:
                    sheet_name = sheet_name.replace(char, '')
                
                # 出力先ファイルパスを生成
                output_file = os.path.join(output_folder, f"{base_filename}_{sheet_name}.pdf")
                
                # シートをPDFに変換（0はPDF形式を意味する）
                sheet.ExportAsFixedFormat(0, output_file)
                print(f"変換成功: {output_file}")
            except Exception as e:
                print(f"シート変換エラー (シート名: {sheet.Name}, ファイル: {excel_path}): {e}")
                traceback.print_exc()
                continue  # 次のシートへ処理を継続
    except Exception as e:
        print(f"シート走査中のエラー ({excel_path}): {e}")
    finally:
        # ワークブックを閉じ、Excelアプリケーションを終了
        workbook.Close(False)
        excel.Quit()

def main():
    # 出力フォルダが存在しない場合は作成する
    if not os.path.exists(OUTPUT_FOLDER):
        try:
            os.makedirs(OUTPUT_FOLDER)
        except Exception as e:
            print(f"出力フォルダ作成エラー ({OUTPUT_FOLDER}): {e}")
            return

    # 指定フォルダ内のサブディレクトリも含む全Excelファイルを走査
    for root, dirs, files in os.walk(SOURCE_FOLDER):
        for file in files:
            if file.lower().endswith(".xlsx"):
                excel_file_path = os.path.join(root, file)
                try:
                    print(f"処理開始: {excel_file_path}")
                    convert_excel_to_pdf(excel_file_path, OUTPUT_FOLDER)
                except Exception as e:
                    print(f"ファイル処理中エラー ({excel_file_path}): {e}")
                    traceback.print_exc()
                    continue

if __name__ == "__main__":
    main()
