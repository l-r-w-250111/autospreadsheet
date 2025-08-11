# -*- coding: utf-8 -*-
"""
設定ファイル
"""
import os

# --- LibreOffice関連の設定 ---
# ご自身の環境に合わせてLibreOfficeのインストールパスを指定してください
LO_PATH = r"C:\Program Files\LibreOffice\program"

# UNO接続文字列
UNO_CONNECTION_STRING = "uno:socket,host=localhost,port=2002;urp;"

# --- LLM関連の設定 ---
# OllamaのAPIエンドポイント
OLLAMA_API_URL = "http://localhost:11434/api/generate"

# コード生成に使用するモデル名
CODE_GENERATOR_MODEL = "gemma3:12b"

# 画像検証に使用するマルチモーダルモデル名
IMAGE_VERIFIER_MODEL = "gemma3:12b"

# --- スクリプト内部で使用するパス ---
LO_PYTHON_PATH = os.path.join(LO_PATH, "python-core", "lib")
LIBREOFFICE_EXECUTABLE = os.path.join(LO_PATH, "scalc.exe")
