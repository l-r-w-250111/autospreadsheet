import subprocess
import time
import os
import uno

UNO_CONNECTION_STRING = "uno:socket,host=localhost,port=2002;urp;"
LIBREOFFICE_PATH = "C:\Program Files\LibreOffice\program\scalce.exe" 

def check_libreoffice_connection(retries=5, delay=5):
    """
    Checks if a LibreOffice process is already running and listening on the UNO port.
    If not, it attempts to start it and retries connecting.
    """
    for i in range(retries):
        try:
            localContext = uno.getComponentContext()
            resolver = localContext.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", localContext)
            ctx = resolver.resolve(UNO_CONNECTION_STRING + "StarOffice.ComponentContext")
            print("LibreOffice is running and connected via UNO.")
            return True
        except Exception:
            if i == 0:
                print("LibreOffice is not running or not connected. Attempting to start it...")
                try:
                    subprocess.Popen([LIBREOFFICE_PATH, f'--accept={UNO_CONNECTION_STRING}', '--norestore'])
                    print("Waiting for LibreOffice to start...")
                except FileNotFoundError:
                    print(f"Error: Could not find LibreOffice executable at {LIBREOFFICE_PATH}")
                    print("Please ensure LibreOffice is installed in the correct directory or update the path in libreoffice_manager.py.")
                    return False
                except Exception as e:
                    print(f"Failed to start LibreOffice: {e}")
                    return False
            
            print(f"Connection attempt {i+1}/{retries} failed. Retrying in {delay} seconds...")
            time.sleep(delay)

    print("Failed to connect to LibreOffice after multiple retries.")
    return False

def get_libreoffice_context():
    """
    LibreOfficeのUNOコンテキスト、デスクトップ、現在のドキュメントを取得する。
    """
    try:
        localContext = uno.getComponentContext()
        resolver = localContext.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", localContext)
        ctx = resolver.resolve(UNO_CONNECTION_STRING + "StarOffice.ComponentContext")
        smgr = ctx.ServiceManager
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        doc = desktop.getCurrentComponent()
        return ctx, desktop, doc
    except Exception as e:
        print(f"LibreOfficeコンテキストの取得に失敗しました: {e}")
        return None, None, None

def stop_libreoffice(process):
    """
    This function does nothing as LibreOffice is expected to be managed manually.
    """
    pass

def set_cell_value(sheet, cell_address, value):
    """
    指定されたシートのセルに値を設定する。
    cell_address: 例 "A1"
    """
    try:
        cell = sheet.getCellRangeByName(cell_address)
        if isinstance(value, (int, float)):
            cell.setValue(float(value))
        else:
            cell.setString(str(value))
        print(f"セル {cell_address} に '{value}' を設定しました。")
    except Exception as e:
        print(f"セル {cell_address} への値設定中にエラーが発生しました: {e}")

def get_cell_value(sheet, cell_address):
    """
    指定されたシートのセルの値を取得する。
    cell_address: 例 "A1"
    """
    try:
        cell = sheet.getCellRangeByName(cell_address)
        if cell.getType() == uno.com.sun.star.table.CellContentType.VALUE:
            return cell.getValue()
        else:
            return cell.getString()
    except Exception as e:
        print(f"セル {cell_address} の値取得中にエラーが発生しました: {e}")
        return None

def get_sheet(doc, sheet_name):
    """
    指定された名前のシートを取得する。
    """
    try:
        sheets = doc.getSheets()
        if sheets.hasByName(sheet_name):
            return sheets.getByName(sheet_name)
        else:
            print(f"シート '{sheet_name}' が見つかりません。")
            return None
    except Exception as e:
        print(f"シート '{sheet_name}' の取得中にエラーが発生しました: {e}")
        return None

def save_document(doc, file_path):
    """
    ドキュメントを保存する。
    """
    try:
        file_url = uno.systemPathToFileUrl(os.path.abspath(file_path))
        doc.storeToURL(file_url, ())
        print(f"ドキュメントを保存しました: {file_path}")
    except Exception as e:
        print(f"ドキュメントの保存中にエラーが発生しました: {e}")

def close_document(doc):
    """
    ドキュメントを閉じる。
    """
    try:
        doc.dispose()
        print("ドキュメントを閉じました。")
    except Exception as e:
        print(f"ドキュメントのクローズ中にエラーが発生しました: {e}")

def close_libreoffice(doc, desktop):
    """
    LibreOfficeアプリケーションを閉じる。
    """
    try:
        if doc:
            doc.dispose()
        if desktop:
            desktop.terminate()
        print("LibreOfficeを閉じました。")
    except Exception as e:
        print(f"LibreOfficeの終了中にエラーが発生しました: {e}")


if __name__ == "__main__":
    if check_libreoffice_connection():
        print("Connection successful. You can now interact with LibreOffice via UNO.")
    else:
        print("Connection failed. Please follow the instructions above to start LibreOffice.")
