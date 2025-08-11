import re
import sys
import os
import json
import uno
from llm_wrapper import invoke_llm, GENERATOR_PROMPT_TEMPLATE
from executor import execute_and_verify
from libreoffice_manager import check_libreoffice_connection
from config import IMAGE_VERIFIER_MODEL

def extract_code_and_query(response_text):
    """
    LLMの応答からPythonコードと検証用JSONクエリを抽出する。
    """
    # Pythonコードの抽出
    code_pattern = re.compile(r"```python\n(.*?)\n```", re.DOTALL)
    code_match = code_pattern.search(response_text)
    code = code_match.group(1).strip() if code_match else None

    # JSONクエリの抽出
    json_pattern = re.compile(r"```json\n(.*?)\n```", re.DOTALL)
    json_match = json_pattern.search(response_text)
    query_str = json_match.group(1).strip() if json_match else "{}"
    
    try:
        query = json.loads(query_str)
    except json.JSONDecodeError:
        query = {}

    return code, query

def main():
    """
    メインの自己改善ループを実行する。
    """
    instruction = input("どのような操作をしますか？（例: A1セルに'Hello'と入力）: ")
    if not instruction.strip():
        print("指示が入力されなかったため、処理を終了します。")
        return

    max_iterations = 5
    feedback_history = "なし"
    final_code = ""

    print(f"--- 初期指示 ---\n{instruction}\n")

    if not check_libreoffice_connection():
        return

    try:
        local_context = uno.getComponentContext()
        resolver = local_context.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local_context)
        context = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
        desktop = context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", context)
        doc = desktop.getCurrentComponent()
    except Exception as e:
        print(f"LibreOfficeへの接続に失敗しました: {e}")
        return

    try:
        for current_iteration in range(1, max_iterations + 1):
            print(f"--- イテレーション {current_iteration}/{max_iterations} ---")

            print("1. コードと検証クエリを生成中...")
            prompt = GENERATOR_PROMPT_TEMPLATE.format(
                instruction=instruction,
                feedback_history=feedback_history
            )
            generated_text = invoke_llm(prompt)
            if not generated_text:
                print("コード生成に失敗しました。処理を中断します。")
                break

            code_to_execute, verification_query = extract_code_and_query(generated_text)

            if not code_to_execute:
                print("応答からPythonコードを抽出できませんでした。")
                feedback_history += f"\n試行{current_iteration}: コードブロックが生成されませんでした。"
                continue

            print(f"生成されたコード:\n---\n{code_to_execute}\n---")
            print(f"生成された検証クエリ:\n---\n{json.dumps(verification_query, indent=2, ensure_ascii=False)}\n---")

            print("2. コードを実行し、ハイブリッド検証中...")
            verification_result, is_pass = execute_and_verify(
                code_string=code_to_execute,
                verification_query=verification_query,
                doc=doc,
                desktop=desktop,
                instruction=instruction,
                image_verifier_model=IMAGE_VERIFIER_MODEL
            )

            print(f"検証結果:\n---\n{verification_result}\n---")

            if is_pass:
                print("\n--- タスク成功！ ---")
                final_code = code_to_execute
                break
            else:
                print("\n--- 失敗。フィードバックを次の試行に活かします。 ---")
                feedback_history += f"\n# 試行 {current_iteration}: 失敗\n"
                feedback_history += f"コード:\n{code_to_execute}\n"
                feedback_history += f"判定結果:\n{verification_result}\n"

            if current_iteration == max_iterations:
                print("\n--- 最大試行回数に達しました ---")

    finally:
        pass

    print("\n--- 処理完了 ---")
    if final_code:
        print(f"最終的に成功したコード:\n{final_code}")
    else:
        print("タスクは指定された試行回数内に成功しませんでした。")

if __name__ == "__main__":
    main()
