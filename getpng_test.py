import uno
import os
from com.sun.star.beans import PropertyValue

def export_active_sheet_to_png(doc, output_path):
    """
    Calcドキュメントのアクティブなシートを1ページに収まるように調整し、PNGファイルとしてエクスポートします。
    """
    try:
        controller = doc.getCurrentController()
        sheet = controller.getActiveSheet()

        # --- 印刷範囲の計算ロジック ---
        
        # 1a. データが入力されているセルの範囲を基準に初期の最大行・列を設定
        cursor = sheet.createCursor()
        cursor.gotoEndOfUsedArea(False)
        data_range_address = cursor.getRangeAddress()
        max_row = data_range_address.EndRow
        max_col = data_range_address.EndColumn

        # 1b. 図形オブジェクトの実際の表示サイズと位置から最大範囲を計算
        draw_page = sheet.getDrawPage()
        if draw_page and draw_page.hasElements():
            columns = sheet.getColumns()
            rows = sheet.getRows()

            for i in range(draw_page.getCount()):
                shape = draw_page.getByIndex(i)
                
                # サイズや位置が取得できないオブジェクトはスキップ
                if not hasattr(shape, 'getSize') or not hasattr(shape, 'getPosition'):
                    continue

                shape_size = shape.getSize()
                shape_pos = shape.getPosition()  # アンカーからの相対位置
                anchor = shape.getAnchor()
                
                # アンカーの開始セルを特定
                start_col_idx, start_row_idx = 0, 0
                if hasattr(anchor, 'getRangeAddress'):
                    anchor_range = anchor.getRangeAddress()
                    start_col_idx = anchor_range.StartColumn
                    start_row_idx = anchor_range.StartRow
                elif hasattr(anchor, 'getCellAddress'):
                    anchor_cell = anchor.getCellAddress()
                    start_col_idx = anchor_cell.Column
                    start_row_idx = anchor_cell.Row

                # アンカーセルの絶対座標 (単位: 1/100mm) を取得
                anchor_abs_pos_x = columns.getByIndex(start_col_idx).Position
                anchor_abs_pos_y = rows.getByIndex(start_row_idx).Position
                
                # オブジェクトの右下端の絶対座標を計算
                shape_end_x = int(anchor_abs_pos_x.X) + int(shape_pos.X) + int(shape_size.Width)
                shape_end_y = int(anchor_abs_pos_y.Y) + int(shape_pos.Y) + int(shape_size.Height)

                # 「セルウォーク」で右下端が含まれるセルを特定
                end_col_idx = start_col_idx
                current_x = anchor_abs_pos_x
                for c in range(start_col_idx, columns.getCount()):
                    col = columns.getByIndex(c)
                    if not col.IsVisible: continue
                    # Note: Position of the *next* column is the end of the current one
                    if c + 1 < columns.getCount():
                        next_col_pos = columns.getByIndex(c + 1).Position
                        if next_col_pos.X >= shape_end_x:
                            end_col_idx = c
                            break
                    else: # Last column
                        end_col_idx = c
                        break

                end_row_idx = start_row_idx
                current_y = anchor_abs_pos_y
                for r in range(start_row_idx, rows.getCount()):
                    row = rows.getByIndex(r)
                    if not row.IsVisible: continue
                    if r + 1 < rows.getCount():
                        next_row_pos = rows.getByIndex(r + 1).Position
                        if next_row_pos.Y >= shape_end_y:
                            end_row_idx = r
                            break
                    else: # Last row
                        end_row_idx = r
                        break
                
                # 最大行・列を更新
                if end_row_idx > max_row:
                    max_row = end_row_idx
                if end_col_idx > max_col:
                    max_col = end_col_idx

        # 1c. A1セルから計算された最終的な右下端までを印刷範囲として設定
        print_area = sheet.getCellRangeByPosition(0, 0, max_col, max_row).getRangeAddress()
        sheet.setPrintAreas((print_area,))

        # 2. ページスタイルを1ページにスケールするように設定
        # 注意: この処理はドキュメントのページスタイルを変更します。
        style_name = sheet.PageStyle
        style_families = doc.getStyleFamilies()
        page_styles = style_families.getByName("PageStyles")
        page_style = page_styles.getByName(style_name)
        page_style.ScaleToPagesX = 1
        page_style.ScaleToPagesY = 1

        # 3. 'calc_png_Export' フィルターを使用してシートをエクスポート
        # このフィルターはアクティブなシートの印刷範囲をエクスポートします。
        output_url = uno.systemPathToFileUrl(output_path)
        
        filter_data = (
            PropertyValue("FilterName", 0, "calc_png_Export", 0),
            # PixelWidth/Heightを指定しないことで、LibreOfficeが
            # シートの内容と印刷設定に基づいてサイズを自動決定します。
        )

        doc.storeToURL(output_url, filter_data)
        print(f"シート '{sheet.getName()}' が {output_path} にエクスポートされました。")

    except Exception as e:
        print(f"エクスポート中にエラーが発生しました: {e}")


def main():
    """
    LibreOfficeに接続し、現在のCalcドキュメントを取得して、
    アクティブなシートをPNGファイルにエクスポートします。
    """
    try:
        local_ctx = uno.getComponentContext()
        resolver = local_ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local_ctx
        )
        ctx = resolver.resolve(
            "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext"
        )
        smgr = ctx.ServiceManager
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

        doc = desktop.getCurrentComponent()
        if not doc:
            print("ドキュメントが開かれていません。")
            return
        
        if not doc.supportsService("com.sun.star.sheet.SpreadsheetDocument"):
            print("現在のドキュメントはCalcスプレッドシートではありません。")
            return

        # 出力パスをカレントディレクトリに設定します
        # os.getcwd() はスクリプトを実行したディレクトリの絶対パスを返します
        current_directory = os.getcwd()
        output_path = os.path.join(current_directory, "sheet.png")
        
        export_active_sheet_to_png(doc, output_path)

    except Exception as e:
        print(f"エラーが発生しました: {e}")

if __name__ == "__main__":
    main()
