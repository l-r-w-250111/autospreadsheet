import uno

def get_calc_state(queries):
    """
    実行中のLibreOffice Calcインスタンスに接続し、指定された複数の情報を取得する。

    Args:
        queries (dict): 取得したい情報のクエリ。
                        例: {"cell_value": "A1", "active_sheet_name": True, "sheet_count": True}

    Returns:
        dict: クエリに対する結果のキーと値のペア。
    """
    results = {}
    try:
        # UNOコンポーネントの取得
        local_context = uno.getComponentContext()
        resolver = local_context.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local_context)
        context = resolver.resolve(
            "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
        desktop = context.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.Desktop", context)
        
        doc = desktop.getCurrentComponent()
        if not hasattr(doc, "Sheets"):
            return {"error": "アクティブなドキュメントがCalcのスプレッドシートではありません。"}

        # クエリに基づいて情報を収集
        if queries.get("cell_values"):
            for cell_address in queries["cell_values"]:
                try:
                    if "." in cell_address:
                        sheet_name, cell = cell_address.split(".", 1)
                        sheet = doc.Sheets.getByName(sheet_name)
                    else:
                        sheet = doc.getCurrentController().getActiveSheet()
                        cell = cell_address
                    
                    cell_range = sheet.getCellRangeByName(cell)
                    # データを行列として取得
                    data_array = cell_range.getDataArray()
                    
                    # 取得したデータを文字列に変換
                    value_str = str(data_array)

                    results[f"cell_values_{cell_address}"] = f"セル範囲 {cell_address} の値: {value_str}"
                except Exception as e:
                    results[f"cell_values_{cell_address}"] = f"セル範囲 {cell_address} の値の取得に失敗: {e}"

        if queries.get("active_sheet_name"):
            try:
                sheet = doc.getCurrentController().getActiveSheet()
                results["active_sheet_name"] = f"アクティブシート名: {sheet.getName()}"
            except Exception as e:
                results["active_sheet_name"] = f"アクティブシート名の取得に失敗: {e}"

        if queries.get("sheet_count"):
            try:
                count = doc.Sheets.getCount()
                results["sheet_count"] = f"シートの総数: {count}"
            except Exception as e:
                results["sheet_count"] = f"シート数の取得に失敗: {e}"

        if queries.get("sheet_names"):
            try:
                names = doc.Sheets.getElementNames()
                results["sheet_names"] = f"全シート名: {list(names)}"
            except Exception as e:
                results["sheet_names"] = f"全シート名の取得に失敗: {e}"

        if queries.get("chart_count"):
            try:
                sheet = doc.getCurrentController().getActiveSheet()
                count = sheet.getCharts().getCount()
                results["chart_count"] = f"アクティブシートのグラフ数: {count}"
            except Exception as e:
                results["chart_count"] = f"グラフ数の取得に失敗: {e}"

        if queries.get("chart_types"):
            try:
                sheet = doc.getCurrentController().getActiveSheet()
                charts = sheet.getCharts()
                chart_info = {}
                for i in range(charts.getCount()):
                    chart_shape = charts.getByIndex(i)
                    chart_doc = chart_shape.getEmbeddedObject()
                    diagram = chart_doc.getDiagram()
                    chart_info[chart_shape.getName()] = diagram.getImplementationName()
                results["chart_types"] = f"各グラフの種類: {chart_info}"
            except Exception as e:
                results["chart_types"] = f"グラフの種類の取得に失敗: {e}"

        if queries.get("document_count"):
            try:
                components = desktop.getComponents()
                # com.sun.star.sheet.SpreadsheetDocument をサポートするコンポーネントを数える
                doc_count = sum(1 for comp in components if hasattr(comp, "Sheets"))
                results["document_count"] = f"Calcドキュメントの総数: {doc_count}"
            except Exception as e:
                results["document_count"] = f"ドキュメント数の取得に失敗: {e}"

    except Exception as e:
        results["error"] = f"Calcの状態取得中にエラーが発生しました: {e}"
    
    return results