import os
import json
import urllib.request
import base64

# Ollamaモデルの設定
OLLAMA_MODEL_NAME = "gemma3:12b" # コード生成に使用するモデル名
OLLAMA_API_URL = "http://localhost:11434/api/generate"

# --- プロンプトテンプレート ---

GENERATOR_PROMPT_TEMPLATE = """
あなたは、ユーザーの指示をPythonのUNO (Universal Network Objects) APIを使ったLibreOffice Calc操作コードに変換するエキスパートです。

# 厳格なルール
- **絶対に**新しいドキュメントを作成してはいけません。`desktop.loadComponentFromURL`の使用は固く禁止します。
- 指示がない限り、常に**既に開かれているアクティブなドキュメント**に対して操作を行ってください。
- 生成するコードは、```python ```で囲んでください。
- **必ず**、生成するコードの冒頭で、以下の定型コードを使ってLibreOfficeに接続し、`doc`と`sheet`オブジェクトを取得してください。
```python
import uno

local_context = uno.getComponentContext()
resolver = local_context.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local_context)
context = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
desktop = context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", context)
doc = desktop.getCurrentComponent()
sheet = doc.getCurrentController().getActiveSheet()
```

# コード生成の注意点
- 生成するコードは、LibreOfficeに付属のPython環境で実行可能な、単独のスクリプトにしてください。
- `doc` と `desktop` オブジェクトが利用可能であることを前提としてください。
- インデントは正確に4スペースで行ってください。

# 参考にするコード例
    # アクティブシートを取得しA1セルに123と入力するコード
        ```python
        import uno

        # LibreOffice に接続
        local_ctx = uno.getComponentContext()
        resolver = local_ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local_ctx
        )
        context = resolver.resolve(
            "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext"
        )
        smgr = context.ServiceManager

        # 現在開いているドキュメントを取得
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", context)
        doc = desktop.getCurrentComponent()

        # スプレッドシートであることを確認
        if not hasattr(doc, "Sheets"):
            raise Exception("スプレッドシートがアクティブではありません。")

        # アクティブシートを取得
        sheet = doc.CurrentController.ActiveSheet

        # A1セルに数値を入力
        cell = sheet.getCellRangeByName("A1")
        cell.setValue(123)  # 数値を入力

        print("A1に 123 を入力しました。")
        ```

    # アクティブなシートの最初のグラフを折れ線グラフに変更する。
        ```python
        import uno
        
        try:
            # LibreOffice に接続
            local_context = uno.getComponentContext()
            resolver = local_context.ServiceManager.createInstanceWithContext(
                "com.sun.star.bridge.UnoUrlResolver", local_context
            )
            context = resolver.resolve(
                "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext"
            )
            desktop = context.ServiceManager.createInstanceWithContext(
                "com.sun.star.frame.Desktop", context
            )

            # Calc ドキュメントを取得
            doc = desktop.getCurrentComponent()
            if not doc.supportsService("com.sun.star.sheet.SpreadsheetDocument"):
                print("エラー: 現在のドキュメントは Calc ではありません。")
            sheet = doc.getCurrentController().getActiveSheet()

            # グラフコレクションを取得
            charts = sheet.getCharts()
            chart_names = charts.getElementNames()
            if not chart_names:
                print("エラー: シートにグラフが存在しません。")
            
            # 最初のグラフを取得
            chart_name_var = chart_names[0]
            chart_shape = charts.getByName(chart_name_var)
            chart_doc = chart_shape.getEmbeddedObject()

            # Diagram を取得して種類を変更
            diagram = chart_doc.getDiagram()
            if diagram.supportsService("com.sun.star.chart.LineDiagram"):
                print("すでに折れ線グラフです。")
            else:
                new_diagram = chart_doc.createInstance("com.sun.star.chart.LineDiagram")
                chart_doc.setDiagram(new_diagram)
                print("グラフ '{{}}' を折れ線グラフに変更しました。".format(chart_name_var))

        except Exception as e:
            print("エラーが発生しました: {{}}".format(e))
        ```

    # アクティブシートに散布図を作成する。
        ```python
        import uno
        from com.sun.star.awt import Rectangle

        try:
            # --- UNO接続 ---
            local_context = uno.getComponentContext()
            resolver = local_context.ServiceManager.createInstanceWithContext(
                "com.sun.star.bridge.UnoUrlResolver", local_context)
            context = resolver.resolve(
                "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
            desktop = context.ServiceManager.createInstanceWithContext(
                "com.sun.star.frame.Desktop", context)
            doc = desktop.getCurrentComponent()
            
            sheet = doc.getCurrentController().getActiveSheet()
            charts = sheet.getCharts()

            # --- グラフの位置とサイズを定義 ---
            rect = Rectangle(10000, 1000, 15000, 8000)

            # --- データ範囲のアドレスを定義 (A1:B6) ---
            data_range = sheet.getCellRangeByName("A1:B6")
            range_address = data_range.getRangeAddress()

            # --- チャートの追加 ---
            charts.addNewByName("SampleScatterChart", rect, (range_address,), True, False)
            
            # --- チャートの種類を散布図に設定 ---
            table_chart = charts.getByName("SampleScatterChart")
            chart_doc = table_chart.getEmbeddedObject()
            diagram = chart_doc.createInstance("com.sun.star.chart.XYDiagram")
            chart_doc.setDiagram(diagram)
            
            print("散布図の作成が完了しました。")

        except Exception as e:
            import traceback
            print("エラーが発生しました: {{}}".format(e))
            print(traceback.format_exc())
        ```

    # アクティブシートの最初のグラフにタイトルを作成する。    
        ```python
        import uno

        # LibreOffice に接続
        local_context = uno.getComponentContext()
        resolver = local_context.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local_context
        )
        context = resolver.resolve(
            "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext"
        )

        desktop = context.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.Desktop", context
        )

        doc = desktop.getCurrentComponent()
        sheet = doc.CurrentController.ActiveSheet

        charts = sheet.Charts
        if charts.getCount() == 0:
            raise Exception("このシートにグラフが見つかりません。")

        chart_name = charts.getByIndex(0).Name
        chart_doc = charts.getByName(chart_name).EmbeddedObject

        if chart_doc.supportsService("com.sun.star.chart.ChartDocument"):
            print("Chart1 形式のグラフを検出")

            chart_doc.HasMainTitle = True  # タイトル表示ON
            title = chart_doc.getTitle()
            title.String = "サンプルグラフタイトル"

            diag = chart_doc.Diagram
            diag.HasXAxisTitle = True  # X軸タイトル表示ON
            diag.XAxisTitle.String = "X軸タイトル"
            diag.HasYAxisTitle = True  # Y軸タイトル表示ON
            diag.YAxisTitle.String = "Y軸タイトル"

            # 変更を反映させるためDiagramを再セット
            chart_doc.Diagram = diag

            # ドキュメント再計算
            doc.calculate()

            # ビュー強制リフレッシュ
            doc.CurrentController.Frame.ContainerWindow.invalidate()
            doc.CurrentController.Frame.ContainerWindow.validate()


        elif chart_doc.supportsService("com.sun.star.chart2.ChartDocument"):
            print("Chart2 形式のグラフを検出")

            title_obj = chart_doc.getTitle()
            if title_obj is None:
                title_obj = chart_doc.createInstance("com.sun.star.chart2.Title")
                chart_doc.setTitle(title_obj)
            title_obj.String = "サンプルグラフタイトル"
            title_obj.IsVisible = True  # タイトル表示ON

            diagram = chart_doc.getFirstDiagram()
            coord_systems = diagram.getCoordinateSystems()
            if coord_systems:
                coord_system = coord_systems[0]

                x_axis = coord_system.getAxisByDimension(0, 0)
                if x_axis:
                    x_title_obj = x_axis.getTitle()
                    if x_title_obj is None:
                        x_title_obj = chart_doc.createInstance("com.sun.star.chart2.Title")
                        x_axis.setTitle(x_title_obj)
                    x_title_obj.String = "X軸タイトル"
                    x_title_obj.IsVisible = True  # X軸タイトル表示ON

                y_axis = coord_system.getAxisByDimension(1, 0)
                if y_axis:
                    y_title_obj = y_axis.getTitle()
                    if y_title_obj is None:
                        y_title_obj = chart_doc.createInstance("com.sun.star.chart2.Title")
                        y_axis.setTitle(y_title_obj)
                    y_title_obj.String = "Y軸タイトル"
                    y_title_obj.IsVisible = True  # Y軸タイトル表示ON

        else:
            raise Exception("未知のグラフ形式です")

        print("グラフタイトル・軸タイトルを設定し、表示をONにしました。")
        ```




# 指示
{instruction}

# 過去の試行と評価 (フィードバック)
{feedback_history}

# あなたが生成するべきPythonコード:
"""

def _image_to_base64(image_path):
    """画像をbase64エンコードする内部ヘルパー関数"""
    if not os.path.exists(image_path):
        return None
    with open(image_path, "rb") as f:
        return base64.b64encode(f.read()).decode("utf-8")

def invoke_llm(prompt):
    """
    指定されたプロンプトを使用してOllama APIを直接呼び出し、応答を返す。
    """
    try:
        data = {
            "model": OLLAMA_MODEL_NAME,
            "prompt": prompt,
            "stream": False
        }
        json_data = json.dumps(data).encode('utf-8')

        req = urllib.request.Request(
            OLLAMA_API_URL,
            data=json_data,
            headers={'Content-Type': 'application/json'}
        )

        with urllib.request.urlopen(req) as response:
            response_text = response.read().decode('utf-8')
            json_response = json.loads(response_text)
            return json_response.get('response', '')

    except Exception as e:
        print("LLMの呼び出し中にエラーが発生しました: {{}}".format(e))
        return None

def invoke_llm_with_image(prompt, image_path, model_name):
    """
    プロンプトと画像をOllamaに送信し、応答を返す。
    画像解析が可能なマルチモーダルモデル（例: llava）を指定してください。
    """
    image_b64 = _image_to_base64(image_path)
    if not image_b64:
        print("エラー: 画像ファイルが見つからないか、読み込めません: {{}}".format(image_path))
        return None

    try:
        data = {
            "model": model_name,
            "prompt": prompt,
            "images": [image_b64],
            "stream": False
        }
        json_data = json.dumps(data).encode('utf-8')

        req = urllib.request.Request(
            OLLAMA_API_URL,
            data=json_data,
            headers={'Content-Type': 'application/json'}
        )

        with urllib.request.urlopen(req) as response:
            response_text = response.read().decode('utf-8')
            json_response = json.loads(response_text)
            return json_response.get('response', '')

    except Exception as e:
        print("画像付きLLMの呼び出し中にエラーが発生しました: {{}}".format(e))
        return None