import os
import traceback
import pandas as pd

# グローバル変数として、元フォルダと出力フォルダのパスを定義
SOURCE_FOLDER = r"input"
OUTPUT_FOLDER = r"output_md"

def convert_excel_to_md(excel_path, output_folder):
    """
    指定したExcelファイル内の全シートをMarkdown形式に変換して保存する関数。
    Markdownのファイル名は "{Excelファイルのファイル名}_{シート名}.md" 形式。
    """
    abs_excel_path = os.path.abspath(excel_path)
    if not os.path.exists(abs_excel_path):
        print(f"ファイルが存在しません: {abs_excel_path}")
        return

    try:
        # Excel ファイルを読み込み（全シートを辞書で取得）
        xls = pd.ExcelFile(abs_excel_path)
        sheets = xls.sheet_names
    except Exception as e:
        print(f"Excel 読み込みエラー: {e}")
        traceback.print_exc()
        return

    base_filename = os.path.splitext(os.path.basename(excel_path))[0]
    for sheet_name in sheets:
        try:
            # シートをデータフレームとして読み込み
            df = pd.read_excel(xls, sheet_name=sheet_name)
            
            # ファイル名に使えない文字を除去
            safe_sheet = sheet_name
            invalid_chars = '<>:"/\\|?*'
            for ch in invalid_chars:
                safe_sheet = safe_sheet.replace(ch, '')

            # 出力先ファイルパス生成
            output_file = os.path.join(output_folder, f"{base_filename}_{safe_sheet}.md")
            
            # DataFrame を Markdown 化してファイルに書き込む
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(df.to_markdown(index=False))
            
            print(f"変換成功: {output_file}")
        except Exception as e:
            print(f"シート変換エラー (シート: {sheet_name}, ファイル: {excel_path}): {e}")
            traceback.print_exc()
            continue

def main():
    # 出力フォルダの準備
    abs_output_folder = os.path.abspath(OUTPUT_FOLDER)
    if not os.path.exists(abs_output_folder):
        try:
            os.makedirs(abs_output_folder)
        except Exception as e:
            print(f"出力フォルダ作成エラー ({abs_output_folder}): {e}")
            return

    # input フォルダ以下の全 Excel ファイルを再帰的に処理
    for root, dirs, files in os.walk(SOURCE_FOLDER):
        for file in files:
            if file.lower().endswith(".xlsx"):
                excel_file_path = os.path.join(root, file)
                print(f"処理開始: {excel_file_path}")
                try:
                    convert_excel_to_md(excel_file_path, abs_output_folder)
                except Exception as e:
                    print(f"ファイル処理中エラー ({excel_file_path}): {e}")
                    traceback.print_exc()
                    continue

if __name__ == "__main__":
    main()
