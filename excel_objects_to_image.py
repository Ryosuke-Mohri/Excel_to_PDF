import os
import traceback
import win32com.client

# グローバル変数として、元フォルダと出力フォルダのパスを定義
SOURCE_FOLDER = r"input"
OUTPUT_FOLDER = r"output_images"

def export_sheet_objects_as_image(sheet, base_filename, output_folder):
    """
    シート上のすべてのオブジェクト（図形、テキストボックス、画像、グラフなど）を
    順序関係を保ったまま一枚の画像として出力する。
    """
    # シェイプ／チャートオブジェクト名を収集
    shape_names = [sh.Name for sh in sheet.Shapes]
    chart_names = [ch.Name for ch in sheet.ChartObjects()]
    all_names = shape_names + chart_names

    if not all_names:
        print(f"  シート '{sheet.Name}' にオブジェクトが見つかりませんでした。")
        return

    # 安全なシート名
    safe_sheet = sheet.Name
    for ch in '<>:"/\\|?*':
        safe_sheet = safe_sheet.replace(ch, '')

    try:
        # すべてのオブジェクトをまとめてコピー（セルは無視）
        shapes_range = sheet.Shapes.Range(all_names)
        shapes_range.CopyPicture(Appearance=1, Format=2)  # 1=画質優先, 2=ビットマップ

        # 一時チャートをシート上に作成して、コピーした画像を貼り付け
        width = sheet.UsedRange.Width
        height = sheet.UsedRange.Height
        chart_obj = sheet.ChartObjects().Add(Left=0, Top=0, Width=width, Height=height)
        chart = chart_obj.Chart
        chart.Paste()

        # エクスポート先パス
        output_path = os.path.join(
            output_folder,
            f"{base_filename}_{safe_sheet}.png"
        )

        # 画像として保存
        chart.Export(output_path)
        print(f"  画像出力成功: {output_path}")

        # 一時チャートを削除
        chart_obj.Delete()

    except Exception as e:
        print(f"  オブジェクト抽出エラー (シート: {sheet.Name}): {e}")
        traceback.print_exc()

def main():
    # Excel アプリケーション起動
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    # 出力フォルダ準備
    abs_out = os.path.abspath(OUTPUT_FOLDER)
    if not os.path.exists(abs_out):
        os.makedirs(abs_out, exist_ok=True)

    # input フォルダ内を再帰的に探索
    for root, _, files in os.walk(SOURCE_FOLDER):
        for file in files:
            if not file.lower().endswith(".xlsx"):
                continue

            excel_path = os.path.join(root, file)
            base_filename = os.path.splitext(os.path.basename(file))[0]
            print(f"処理開始: {excel_path}")

            try:
                wb = excel.Workbooks.Open(os.path.abspath(excel_path))
            except Exception as e:
                print(f"  ワークブックオープンエラー: {e}")
                continue

            # 各シートごとにオブジェクトを画像化
            for sheet in wb.Worksheets:
                export_sheet_objects_as_image(sheet, base_filename, abs_out)

            # ワークブックを閉じる
            wb.Close(False)

    # Excel アプリケーション終了
    excel.Quit()

if __name__ == "__main__":
    main()
