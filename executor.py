import sys
import os

# LibreOffice UNOモジュールへのパスを動的に追加
# LibreOfficeのインストールパスに合わせて適宜変更してください
LO_PATH = r"C:\Program Files\LibreOffice\program"
LO_PYTHON_PATH = os.path.join(LO_PATH, "python-core", "lib")

if LO_PATH not in sys.path:
    sys.path.insert(0, LO_PATH)
if LO_PYTHON_PATH not in sys.path:
    sys.path.insert(0, LO_PYTHON_PATH)

import uno
import traceback
from com.sun.star.beans import PropertyValue
from libreoffice_manager import set_cell_value, get_cell_value, get_sheet, save_document, close_document
from llm_wrapper import invoke_llm_with_image
import capture_png

# 画像判定用のプロンプトテンプレート
VERIFICATION_PROMPT_TEMPLATE = """
以下は、表計算ソフトLibreOffice Calcに対する操作指示と、その操作が実行された後のシートの画像です。
あなたのタスクは、画像を見て、指示された操作が正しく実行されているかを厳密に判定することです。

# 指示内容
{instruction}

# 判定
- 指示通りの変更が正確に行われていますか？
- 指示にない余計な変更（副作用）はありませんか？

判定結果を、必ず以下の形式で、理由を簡潔に述べてください。

判定: [PASSまたはFAIL]
理由: [判定の根拠]
"""

def execute_code(code_string, doc, desktop): # doc, desktop を引数に追加
    """
    与えられたPythonコードを実行する。
    LibreOfficeのUNO環境で実行されることを想定。
    """
    try:
        # 生成されたコードを前処理: uno.awt.Rectangle を uno.createUnoStruct に置換
        processed_code_string = code_string.replace(
            "uno.awt.Rectangle(", "uno.createUnoStruct(\"com.sun.star.awt.Rectangle\", "
        )

        # Provide doc and desktop to the execution context
        # libreoffice_managerの関数も渡す
        exec(processed_code_string, {
            'doc': doc,
            'desktop': desktop,
            'set_cell_value': set_cell_value,
            'get_cell_value': get_cell_value,
            'get_sheet': get_sheet,
            'save_document': save_document,
            'close_document': close_document
        })
        return None, "コードは正常に実行されました。"
    except Exception as e:
        # エラーが発生した場合、その情報を返す
        error_message = f"コード実行エラー: {type(e).__name__}: {e}\n"
        error_message += "".join(traceback.format_exc())
        return error_message, None

def save_sheet_as_png(doc, output_path):
    """
    現在のLibreOffice CalcのアクティブシートをPNG画像として保存する。
    capture_png.pyの機能を利用する。
    """
    try:
        capture_png.export_active_sheet_to_png(doc, output_path)
        return None
    except Exception as e:
        error_message = f"PNG保存エラー: {type(e).__name__}: {e}\n"
        error_message += "".join(traceback.format_exc())
        print(error_message)
        return error_message

def execute_and_verify_with_image(code_string, doc, desktop, instruction, image_verifier_model):
    """
    コードを実行し、結果をPNG画像で保存し、LLMで成否を判定する。
    """
    # 1. コードを実行
    execution_error, result_message = execute_code(code_string, doc, desktop)
    if execution_error:
        return f"実行時エラー: {execution_error}", False

    # 2. 結果をPNGに保存
    temp_image_path = os.path.join(os.getcwd(), "verification.png")
    save_error = save_sheet_as_png(doc, temp_image_path)
    if save_error:
        return f"画像保存エラー: {save_error}", False

    # 3. LLMによる画像判定
    try:
        prompt = VERIFICATION_PROMPT_TEMPLATE.format(instruction=instruction)
        verification_result = invoke_llm_with_image(
            prompt=prompt,
            image_path=temp_image_path,
            model_name=image_verifier_model
        )

        if verification_result is None:
            return "画像判定LLMからの応答がありません。", False

        # 判定結果からPASS/FAILを抽出（簡易的な抽出）
        is_pass = "pass" in verification_result.lower()
        
        return verification_result, is_pass

    finally:
        # 4. 一時ファイルを削除 (デバッグのためコメントアウト)
        # if os.path.exists(temp_image_path):
        #     os.remove(temp_image_path)
        pass # Do nothing for now
