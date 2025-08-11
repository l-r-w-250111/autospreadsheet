import sys
import os
import json
from config import LO_PATH, LO_PYTHON_PATH

# LibreOffice UNOモジュールへのパスを動的に追加
if LO_PATH not in sys.path:
    sys.path.insert(0, LO_PATH)
if LO_PYTHON_PATH not in sys.path:
    sys.path.insert(0, LO_PYTHON_PATH)

import uno
import traceback
from com.sun.star.beans import PropertyValue
from libreoffice_manager import set_cell_value, get_cell_value, get_sheet, save_document, close_document
from llm_wrapper import invoke_llm_with_image
from state_extractor import get_calc_state
import capture_png

# ハイブリッド検証用の新しいプロンプトテンプレート
VERIFICATION_PROMPT_TEMPLATE = """You are a meticulous and detail-oriented AI assistant for spreadsheet verification.
Your task is to determine if an operation was successful by comparing the user's instruction against objective data from the application's API and a screenshot of the user interface.

# User Instruction
{instruction}

# Objective State Data (from API)
This data is the ground truth. Trust this data over the image if there is a conflict.
```json
{objective_state}
```

# Analysis Steps
1.  **Analyze Objective Data**: Does the objective state data reflect the result of the user's instruction? For example, if the instruction was to create a chart, does `chart_count` show an increase?
2.  **Examine Image**: Look at the screenshot. Does it visually confirm the state described in the objective data?
3.  **Synthesize and Conclude**: Based on the objective data (primary source) and the image (secondary confirmation), was the instruction successfully executed?

**Crucial Instruction**: Do NOT invent or hallucinate elements. If the objective data says `chart_count` is 0, and you think you see a chart in the image, you MUST conclude there is no chart. The objective data is the truth.

# Final Verdict
Verdict: [PASS or FAIL]
Reason: [Your reasoning, referencing both objective data and the image]
"""

def execute_code(code_string, doc, desktop):
    """
    Executes the given Python code string.
    """
    try:
        processed_code_string = code_string.replace(
            "uno.awt.Rectangle(", "uno.createUnoStruct(\"com.sun.star.awt.Rectangle\", "
        )
        exec(processed_code_string, {
            'doc': doc,
            'desktop': desktop,
            'set_cell_value': set_cell_value,
            'get_cell_value': get_cell_value,
            'get_sheet': get_sheet,
            'save_document': save_document,
            'close_document': close_document
        })
        return None, "Code executed successfully."
    except Exception as e:
        error_message = f"Code execution error: {type(e).__name__}: {e}\n"
        error_message += "".join(traceback.format_exc())
        return error_message, None

def save_sheet_as_png(doc, output_path):
    """
    Saves the active sheet as a PNG image.
    """
    try:
        capture_png.export_active_sheet_to_png(doc, output_path)
        return None
    except Exception as e:
        error_message = f"PNG save error: {type(e).__name__}: {e}\n"
        error_message += "".join(traceback.format_exc())
        print(error_message)
        return error_message

def execute_and_verify(code_string, verification_query, doc, desktop, instruction, image_verifier_model):
    """
    Executes code, gets objective state, and verifies the result with an image and state data.
    """
    # 1. Execute the code
    execution_error, result_message = execute_code(code_string, doc, desktop)
    if execution_error:
        return f"Execution Error: {execution_error}", False

    # 2. Get objective state from the application
    try:
        objective_state = get_calc_state(verification_query)
    except Exception as e:
        return f"State Extraction Error: {e}", False

    # 3. Save the resulting state as a PNG image
    temp_image_path = os.path.join(os.getcwd(), "verification.png")
    save_error = save_sheet_as_png(doc, temp_image_path)
    if save_error:
        return f"Image Save Error: {save_error}", False

    # 4. Verify with LLM using both objective data and the image
    try:
        prompt = VERIFICATION_PROMPT_TEMPLATE.format(
            instruction=instruction,
            objective_state=json.dumps(objective_state, indent=2, ensure_ascii=False)
        )
        
        verification_result = invoke_llm_with_image(
            prompt=prompt,
            image_path=temp_image_path,
            model_name=image_verifier_model
        )

        if verification_result is None:
            return "Image verification LLM returned no response.", False

        # Simple pass/fail check
        is_pass = "pass" in verification_result.lower()
        
        return verification_result, is_pass

    finally:
        # Keep the image for debugging purposes
        # if os.path.exists(temp_image_path):
        #     os.remove(temp_image_path)
        pass
