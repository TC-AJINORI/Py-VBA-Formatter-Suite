# active_vba_formatter.py
# v1.0.0 公開
# v1.0.1 二重起動チェック機能
# v1.0.2 公開開始
# v1.0.3 エクセル監視機能を堅牢に。フォーマットした後に1行目に戻るのを修正
# ===================================================================================
#
# Version: 1.0.3
#
# 概要:
#   バックグラウンドで常駐し、アクティブなExcelブックのVBAコードを
#   ファイル保存時に自動でインデント整形するツール。
#   OSのUI言語に応じてメッセージを日本語/英語で表示する。
#
# 依存ライブラリ:
#   pywin32, pystray, Pillow
#
# ===================================================================================

import sys
import time
import os
import win32com.client
import win32gui
import pythoncom
import pywintypes
from pywintypes import com_error
import win32event
import winerror
import win32api
import threading
import pystray
from PIL import Image
from queue import Queue
import ctypes

# --- グローバル変数 ---
stop_event = threading.Event()
log_queue = Queue()

# ===================================================================================
# 1. 多言語メッセージ管理
# ===================================================================================
def is_japanese_os() -> bool:
    """OSのUI言語が日本語であるかどうかを判定する。"""
    try:
        # 日本語のLCIDは1041 (0x0411)
        return ctypes.windll.kernel32.GetUserDefaultUILanguage() == 1041
    except Exception:
        return False

class Messages:
    """UIメッセージをOS言語に応じて管理するクラス。"""
    def __init__(self):
        self.is_jp = is_japanese_os()

    # --- 汎用メッセージ ---
    def app_name(self): return "VBAフォーマッター" if self.is_jp else "VBA Formatter"
    def error_title(self): return "エラー" if self.is_jp else "Error"
    def startup_error_title(self): return "起動エラー" if self.is_jp else "Startup Error"
    def fatal_error_title(self): return "致命的なエラー" if self.is_jp else "Fatal Error"

    # --- ツールチップ関連 ---
    def tooltip_header(self):
        return "Active VBA Formatter(右クリックで終了)\n状態" if self.is_jp else "Active VBA Formatter(Right-click to Exit)\nStatus"
    def tooltip_initializing(self): return "初期化中..." if self.is_jp else "Initializing..."
    def menu_quit(self): return "終了" if self.is_jp else "Exit"

    # --- 監視ループ関連 ---
    def monitoring_started(self): return "監視を開始しました。" if self.is_jp else "Monitoring started."
    def target_switched(self, f): return f"監視対象を切り替え: {f}" if self.is_jp else f"Switched target to: {f}"
    def monitoring_interrupted(self, f): return f"'{f}' の監視を中断しました。" if self.is_jp else f"Stopped monitoring '{f}'."
    def formatting_detected(self, f): return f"'{f}' の保存を検知、整形中..." if self.is_jp else f"Detected save for '{f}', formatting..."
    def formatting_complete(self): return "整形が完了しました。" if self.is_jp else "Formatting complete."
    def connection_lost(self, f): return f"'{f}' との接続が切れました。" if self.is_jp else f"Connection lost with '{f}'."
    def monitoring_vbe(self, f): return f"'{f}' のVBEを監視中..." if self.is_jp else f"Monitoring VBE for '{f}'..."
    def waiting_for_vbe(self, f): return f"'{f}' のVBEが開かれるのを待っています..." if self.is_jp else f"Waiting for VBE of '{f}' to open..."
    def searching_for_book(self): return "アクティブなExcelブックを探しています..." if self.is_jp else "Searching for an active Excel workbook..."
    def waiting_for_excel(self): return "Excelの起動、またはウィンドウの表示を待っています..." if self.is_jp else "Waiting for Excel to start..."
    def monitoring_stopped(self): return "監視を停止しました。" if self.is_jp else "Monitoring stopped."

    # --- ダイアログ関連 ---
    def app_is_running(self): return f"{self.app_name()} は既に起動しています。" if self.is_jp else f"{self.app_name()} is already running."
    def startup_check_error(self, e): return f"二重起動チェック中にエラーが発生しました: {e}" if self.is_jp else f"Error during startup check: {e}"
    def icon_not_found(self, path): return f"アイコンファイルが見つかりません:\n{path}" if self.is_jp else f"Icon file not found:\n{path}"
    def ask_exit_message(self): return "全てのExcelが終了しました。ツールを終了しますか？" if self.is_jp else "All Excel windows have been closed. Exit the tool?"
    def exiting_by_dialog(self): return "ダイアログから終了が選択されました。" if self.is_jp else "Exit selected from dialog."
    def exiting_by_menu(self): return "プログラムを終了しています..." if self.is_jp else "Exiting program..."
    
    # --- 通知用メッセージ ---
    def notification_title(self): return self.app_name()
    def notification_message(self): return "起動しました。VBAコードの自動整形を開始します。" if self.is_jp else "Started. Now monitoring VBA code for auto-formatting."

# グローバルなメッセージインスタンスを作成
messages = Messages()

# ===================================================================================
# 2. コード整形クラス
# ===================================================================================
class VbaFormatter:
    """VBAコードのインデントを自動整形する機能を提供するクラス。"""
    def __init__(self, indent_char: str = "    "):
        self.indent_char = indent_char
        # ... (クラスの内部実装は変更なし)
        self.INDENT_KEYWORDS = ("if", "for", "do", "with", "sub", "public sub", "private sub", "function", "public function", "private function", "property", "public property", "private property", "select case", "type")
        self.DEDENT_KEYWORDS = ("end if", "next", "loop", "end with", "end sub", "end function", "end property", "end select", "end type")
        self.MID_BLOCK_KEYWORDS = ("else", "elseif", "else if")
    def _get_judgement_line(self, code_line: str) -> str:
        clean_line = ""; in_string = False
        for char in code_line:
            if char == '"': in_string = not in_string; continue
            if char == "'" and not in_string: break
            if not in_string: clean_line += char
        return clean_line.strip()
    def format_code(self, code_string: str) -> str:
        lines = code_string.splitlines(); formatted_lines = []; current_indent_level = 0; block_stack = []
        for line in lines:
            stripped_line = line.strip()
            if not stripped_line:
                if formatted_lines and formatted_lines[-1] != "": formatted_lines.append("")
                continue
            judgement_line = self._get_judgement_line(stripped_line.replace("_", "")).lower(); judgement_parts = judgement_line.split()
            first_word = judgement_parts[0] if judgement_parts else ""; first_two_words = " ".join(judgement_parts[:2]) if len(judgement_parts) > 1 else ""
            is_start_block = first_two_words in self.INDENT_KEYWORDS or first_word in self.INDENT_KEYWORDS; is_end_block = first_two_words in self.DEDENT_KEYWORDS or first_word in self.DEDENT_KEYWORDS
            is_mid_block = first_two_words in self.MID_BLOCK_KEYWORDS or first_word in self.MID_BLOCK_KEYWORDS; is_case_statement = first_word == "case" or first_two_words == "case else"
            is_select_case = first_two_words == "select case"; is_end_select = first_two_words == "end select"
            if is_end_select:
                current_indent_level = max(0, current_indent_level - 2)
                if block_stack: block_stack.pop()
            elif is_case_statement:
                if block_stack and block_stack[-1] == 'in_case': current_indent_level = max(0, current_indent_level - 1)
            elif is_mid_block: current_indent_level = max(0, current_indent_level - 1)
            elif is_end_block:
                current_indent_level = max(0, current_indent_level - 1)
                if block_stack: block_stack.pop()
            formatted_lines.append(self.indent_char * current_indent_level + stripped_line)
            is_single_line_if = False
            if first_word == "if" and "then" in judgement_line:
                then_pos = judgement_line.find("then"); rest_of_line = judgement_line[then_pos + 4:].strip()
                if rest_of_line and not rest_of_line.startswith("'"): is_single_line_if = True
            if is_select_case:
                current_indent_level += 1; block_stack.append('select')
            elif is_case_statement:
                current_indent_level += 1
                if block_stack and block_stack[-1] == 'select': block_stack[-1] = 'in_case'
            elif (is_start_block and not is_single_line_if) or is_mid_block:
                current_indent_level += 1
                if is_start_block and not is_single_line_if: block_stack.append('other')
        return "\n".join(formatted_lines)

# ===================================================================================
# 3. メイン監視ループ
# ===================================================================================

def monitoring_loop(icon):
    """バックグラウンドでExcelの動作を監視し、コード整形を実行するメインスレッド。"""
    log_queue.put(messages.monitoring_started())
    formatter = VbaFormatter()
    target_app = target_hwnd = target_filepath = None
    last_modified_time = 0
    was_window_visible = is_any_excel_window_visible()
    last_status_message = ""
    
    pythoncom.CoInitialize()
    while not stop_event.is_set():
        current_status_message = ""

        # ステップ1: まず、ユーザーが操作可能なExcelウィンドウが存在するかを判断する
        is_visible_now = is_any_excel_window_visible()

        # ステップ2: ウィンドウが存在する場合のみ、アクティブなExcelの情報を取得しようと試みる
        if is_visible_now:
            active_app, active_hwnd, active_filepath = get_active_excel_info()

            # ステップ2a: 新しいターゲットを捕捉した場合
            if active_app and active_hwnd != target_hwnd:
                log_queue.put(messages.target_switched(os.path.basename(active_filepath)))
                target_app, target_hwnd, target_filepath = active_app, active_hwnd, active_filepath
                if os.path.exists(target_filepath):
                    last_modified_time = os.path.getmtime(target_filepath)
                last_status_message = ""

            # ステップ2b: 既存のターゲットが無効になったかチェック
            is_target_valid = False
            if target_app:
                try:
                    _ = target_app.Name; is_target_valid = True
                except com_error:
                    is_target_valid = False
            
            if not is_target_valid and target_app:
                log_queue.put(messages.monitoring_interrupted(os.path.basename(target_filepath)))
                target_app = target_hwnd = target_filepath = None; last_status_message = ""

            # ステップ3: 有効なターゲットに対してフォーマット処理を実行
            if target_app:
                try:
                    if target_app.VBE.MainWindow.Visible:
                        current_modified_time = os.path.getmtime(target_filepath)
                        if current_modified_time > last_modified_time:
                            log_queue.put(messages.formatting_detected(os.path.basename(target_filepath)))
                            last_modified_time = current_modified_time
                            vbe = target_app.VBE
                            for component in vbe.ActiveVBProject.VBComponents:
                                if component.CodeModule.CountOfLines > 0:
                                    code_pane = component.CodeModule.CodePane
                                    try:
                                        original_line = code_pane.GetSelection(1, 1, 1, 80)[0]
                                        top_line = code_pane.TopLine
                                    except com_error:
                                        original_line = 1; top_line = 1
                                    original_code = component.CodeModule.Lines(1, component.CodeModule.CountOfLines)
                                    formatted_code = formatter.format_code(original_code)
                                    if original_code != formatted_code:
                                        component.CodeModule.DeleteLines(1, component.CodeModule.CountOfLines)
                                        component.CodeModule.AddFromString(formatted_code)
                                        try:
                                            code_pane.TopLine = top_line
                                            code_pane.SetSelection(original_line, 1, original_line, 1)
                                        except com_error:
                                            pass
                            log_queue.put(messages.formatting_complete())
                            last_status_message = ""
                        current_status_message = messages.monitoring_vbe(os.path.basename(target_filepath))
                    else:
                        current_status_message = messages.waiting_for_vbe(os.path.basename(target_filepath))
                except com_error:
                    log_queue.put(messages.connection_lost(os.path.basename(target_filepath)))
                    target_app = target_hwnd = target_filepath = None
                    last_status_message = ""
            else:
                current_status_message = messages.searching_for_book()

        # ステップ4: ウィンドウが存在しない場合の処理
        else:
            current_status_message = messages.waiting_for_excel()
            if target_app:
                log_queue.put(messages.monitoring_interrupted(os.path.basename(target_filepath)))
                target_app = target_hwnd = target_filepath = None
                last_status_message = ""

        # ステップ5: 終了判定
        if was_window_visible and not is_visible_now:
            if ask_to_exit():
                log_queue.put(messages.exiting_by_dialog())
                stop_event.set()
                break

        # ステップ6: ツールチップと状態の更新
        if current_status_message and current_status_message != last_status_message:
            log_queue.put(current_status_message)
            last_status_message = current_status_message
        
        was_window_visible = is_visible_now
        time.sleep(1)

    pythoncom.CoUninitialize()
    log_queue.put(messages.monitoring_stopped())





# ===================================================================================
# 4. ヘルパー関数群
# ===================================================================================
def ask_to_exit() -> bool:
    """ツールを終了するか確認するダイアログを、最前面に表示する。"""
    
    # --- この関数を呼び出すスレッドを一時的にフォアグラウンドに設定 ---
    # これにより、後続のダイアログが確実にフォーカスを得られるようになる
    try:
        # このAPIは管理者権限で実行されていないと失敗することがあるため、
        # 失敗してもプログラムが落ちないようにtry-exceptで囲む
        ctypes.windll.user32.AllowSetForegroundWindow(-1) # -1 は現在のプロセスIDを意味する
    except Exception as e:
        # 失敗した場合はデバッグ用にコンソールに出力（製品版では不要）
        #print(f"[DEBUG] AllowSetForegroundWindow failed: {e}")
        pass
    
    # --- スタイル定数 ---
    MB_YESNO = 0x00000004
    MB_ICONQUESTION = 0x00000020
    MB_TOPMOST = 0x00040000      # 常に最前面に表示する
    MB_SETFOREGROUND = 0x00010000 # ダイアログをフォアグラウンドウィンドウにする
    
    IDYES = 6
    
    # --- スタイルを組み合わせてMessageBoxを呼び出す ---
    result = ctypes.windll.user32.MessageBoxW(
        None, 
        messages.ask_exit_message(), 
        messages.app_name(), 
        MB_YESNO | MB_ICONQUESTION | MB_TOPMOST | MB_SETFOREGROUND
    )
    return result == IDYES

def get_active_excel_info():
    """
    フォアグラウンドのExcelアプリケーション情報を取得する。
    ウィンドウハンドルとCOMオブジェクトを別々に取得し、それらが一致するかを検証する。
    """
    try:
        # ステップ1: フォアグラウンドのウィンドウハンドルがExcelか確認
        fg_hwnd = win32gui.GetForegroundWindow()
        if not fg_hwnd or win32gui.GetClassName(fg_hwnd) not in ('XLMAIN', 'EXCEL7'):
            return None, None, None

        # ステップ2: COMテーブルからアクティブなExcelオブジェクトを取得
        app = win32com.client.GetActiveObject("Excel.Application")
        
        # ステップ3: アクティブなワークブックが存在するかを安全にチェック
        # getattrを使い、プロパティが存在しない場合はNoneを返すようにする
        active_wb = getattr(app, 'ActiveWorkbook', None)
        if not active_wb:
            return None, None, None

        # ステップ4: ワークブックのフルパスが取得できるかチェック
        app_path = getattr(active_wb, 'FullName', None)
        if not app_path:
            return None, None, None

        # ステップ5: ウィンドウハンドルとCOMオブジェクトが一致するかを検証
        window_title = win32gui.GetWindowText(fg_hwnd)
        if os.path.basename(app_path) in window_title:
            return app, fg_hwnd, app_path
        else:
            return None, None, None

    except (com_error, pywintypes.error, AttributeError):
        # [修正] AttributeErrorも捕捉対象に加え、より安全にする
        return None, None, None



def is_any_excel_window_visible():
    """表示されているExcelのメインウィンドウが1つでも存在するかを返す。"""
    found = False
    def enum_proc(hwnd, lParam):
        nonlocal found
        # [修正] ウィンドウが可視であり、かつ、ウィンドウタイトルが空でないことを条件に加える
        if (win32gui.IsWindowVisible(hwnd) and 
            win32gui.GetWindowText(hwnd) != "" and 
            win32gui.GetClassName(hwnd) in ('XLMAIN', 'EXCEL7')):
            
            found = True
            return False # 1つ見つかったので列挙を停止
        return True
        
    try:
        win32gui.EnumWindows(enum_proc, None)
    except pywintypes.error:
        pass # APIエラーは無視
        
    return found

# ===================================================================================
# 5. メイン実行ブロック
# ===================================================================================
def main():
    """システムトレイアイコンを設定し、アプリケーションのメインループを開始する。"""
    mutex_name = "Global\\ActiveVBAFormatterMutex_A1B2C3D4"; mutex = None
    MB_ICONWARNING = 0x30; MB_ICONERROR = 0x10
    try:
        mutex = win32event.CreateMutex(None, 1, mutex_name)
        if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
            ctypes.windll.user32.MessageBoxW(None, messages.app_is_running(), messages.startup_error_title(), MB_ICONWARNING)
            if mutex: win32api.CloseHandle(mutex)
            return
    except Exception as e:
        ctypes.windll.user32.MessageBoxW(None, messages.startup_check_error(e), messages.fatal_error_title(), MB_ICONERROR)
        if mutex: win32api.CloseHandle(mutex)
        return

    def update_tooltip(icon):
        header = messages.tooltip_header(); current_log = messages.tooltip_initializing()
        while icon.visible:
            new_log = None
            while not log_queue.empty(): new_log = log_queue.get_nowait().strip()
            if new_log: current_log = new_log
            full_tooltip = f"{header}\n{current_log}"
            if icon.title != full_tooltip: icon.title = full_tooltip
            time.sleep(0.2)

    def on_quit(icon, item):
        log_queue.put(messages.exiting_by_menu()); stop_event.set(); icon.stop()

    def setup(icon):
        """アイコン表示直後に実行される初期設定。"""
        icon.visible = True
        # バックグラウンドスレッドを開始
        threading.Thread(target=update_tooltip, args=(icon,), daemon=True).start()
        threading.Thread(target=monitoring_loop, args=(icon,), daemon=True).start()

        # 起動通知を表示する
        icon.notify(
            messages.notification_message(),
            messages.notification_title()
        )

    try:
        base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        icon_path = os.path.join(base_path, "active_vba_formatter.ico")
        image = Image.open(icon_path)
    except FileNotFoundError:
        ctypes.windll.user32.MessageBoxW(None, messages.icon_not_found(icon_path), messages.startup_error_title(), MB_ICONERROR)
        if mutex: win32event.ReleaseMutex(mutex); win32api.CloseHandle(mutex)
        return
        
    initial_tooltip = f"{messages.tooltip_header()}\n{messages.tooltip_initializing()}"
    menu = (pystray.MenuItem(messages.menu_quit(), on_quit),)
    icon = pystray.Icon("vba_formatter", image, initial_tooltip, menu)
    icon.run(setup)
    if mutex: win32event.ReleaseMutex(mutex); win32api.CloseHandle(mutex)

if __name__ == "__main__":
    main()

    '''def monitoring_loop(icon):
    """バックグラウンドでExcelの動作を監視し、コード整形を実行するメインスレッド。"""
    log_queue.put(messages.monitoring_started())
    formatter = VbaFormatter()
    target_app = target_hwnd = target_filepath = None
    last_modified_time = 0
    was_window_visible = is_any_excel_window_visible()
    last_status_message = ""
    pythoncom.CoInitialize()
    while not stop_event.is_set():
        current_status_message = ""
        active_app, active_hwnd, active_filepath = get_active_excel_info()
        if active_app and active_hwnd != target_hwnd:
            log_queue.put(messages.target_switched(os.path.basename(active_filepath)))
            target_app, target_hwnd, target_filepath = active_app, active_hwnd, active_filepath
            if os.path.exists(target_filepath): last_modified_time = os.path.getmtime(target_filepath)
            last_status_message = ""
        is_target_valid = False
        if target_app and target_hwnd and win32gui.IsWindow(target_hwnd):
            try:
                _ = target_app.Name; is_target_valid = True
            except com_error: is_target_valid = False
        if not is_target_valid and target_app:
            log_queue.put(messages.monitoring_interrupted(os.path.basename(target_filepath)))
            target_app = target_hwnd = target_filepath = None; last_status_message = ""
        if target_app:
            try:
                if target_app.VBE.MainWindow.Visible:
                    current_modified_time = os.path.getmtime(target_filepath)
                    if current_modified_time > last_modified_time:
                        log_queue.put(messages.formatting_detected(os.path.basename(target_filepath)))
                        last_modified_time = current_modified_time
                        vbe = target_app.VBE
                        for component in vbe.ActiveVBProject.VBComponents:
                            if component.CodeModule.CountOfLines > 0:
                                original_code = component.CodeModule.Lines(1, component.CodeModule.CountOfLines); formatted_code = formatter.format_code(original_code)
                                if original_code != formatted_code:
                                    component.CodeModule.DeleteLines(1, component.CodeModule.CountOfLines); component.CodeModule.AddFromString(formatted_code)
                        log_queue.put(messages.formatting_complete())
                        last_status_message = ""
                    current_status_message = messages.monitoring_vbe(os.path.basename(target_filepath))
                else: current_status_message = messages.waiting_for_vbe(os.path.basename(target_filepath))
            except Exception as e: # ← "as e" を追加
                print(f"monitoring_loop Error: {e}") # ← この行を追加
                log_queue.put(messages.connection_lost(os.path.basename(target_filepath)))
                target_app = target_hwnd = target_filepath = None; last_status_message = ""
        else:
            if is_any_excel_window_visible(): current_status_message = messages.searching_for_book()
            else: current_status_message = messages.waiting_for_excel()
        if current_status_message and current_status_message != last_status_message:
            log_queue.put(current_status_message); last_status_message = current_status_message
        is_visible_now = is_any_excel_window_visible()
        if was_window_visible and not is_visible_now:
            if ask_to_exit():
                log_queue.put(messages.exiting_by_dialog()); stop_event.set(); icon.menu.items[0](icon); break
        was_window_visible = is_visible_now
        time.sleep(1)
    pythoncom.CoUninitialize()
    log_queue.put(messages.monitoring_stopped())'''

    '''def get_active_excel_info():
    """フォアグラウンドのExcelアプリケーション情報を、堅牢な方法で取得する。"""
    try:
        hwnd = win32gui.GetForegroundWindow()
        if not hwnd or win32gui.GetClassName(hwnd) not in ('XLMAIN', 'EXCEL7'): return None, None, None
        ptr = win32com.client.Dispatch("Accessibility.ACC.Client")
        app = ptr.AccessibleObjectFromWindow(hwnd).Application
        if app and app.ActiveWorkbook and app.ActiveWorkbook.FullName:
            return app, hwnd, app.ActiveWorkbook.FullName
    except (com_error, pywintypes.error) as e:
        print(f"get_active_excel_info Error: {e}") # ← この行を追加
    return None, None, None'''

'''def monitoring_loop(icon):
    """バックグラウンドでExcelの動作を監視し、コード整形を実行するメインスレッド。"""
    #print("[DEBUG] monitoring_loop started") # ★追加
    log_queue.put(messages.monitoring_started())
    formatter = VbaFormatter()
    target_app = target_hwnd = target_filepath = None
    last_modified_time = 0
    was_window_visible = is_any_excel_window_visible()
    last_status_message = ""
    pythoncom.CoInitialize()

    loop_count = 0 # ★追加
    while not stop_event.is_set():
        loop_count += 1 # ★追加
        #print(f"\n--- Loop {loop_count} ---") # ★追加

        current_status_message = ""
        
        # 1. Excel情報取得
        #print("[DEBUG] Calling get_active_excel_info...") # ★追加
        active_app, active_hwnd, active_filepath = get_active_excel_info()
        #print(f"[DEBUG] get_active_excel_info returned: app={active_app is not None}, hwnd={active_hwnd}, path={active_filepath}") # ★追加

        # 2. 監視対象切り替え
        if active_app and active_hwnd != target_hwnd:
            #print(f"[DEBUG] Switching target to {os.path.basename(active_filepath)}") # ★追加
            log_queue.put(messages.target_switched(os.path.basename(active_filepath)))
            # (以下略)
            target_app, target_hwnd, target_filepath = active_app, active_hwnd, active_filepath
            if os.path.exists(target_filepath): last_modified_time = os.path.getmtime(target_filepath)
            last_status_message = ""

        # 3. 監視対象の有効性チェック
        is_target_valid = False
        if target_app and target_hwnd and win32gui.IsWindow(target_hwnd):
            try:
                _ = target_app.Name; is_target_valid = True
            except com_error:
                is_target_valid = False
        
        # --- ▼ ここから修正 ▼ ---
        # ターゲットが無効になった場合（Excelが閉じられた、またはクラッシュした場合）
        if not is_target_valid and target_app:
            log_queue.put(messages.monitoring_interrupted(os.path.basename(target_filepath)))
            
            # [重要] 接続が切れた直後にウィンドウの存在を最終確認する
            time.sleep(0.5) # OSの状態安定を待つ
            if not is_any_excel_window_visible():
                if ask_to_exit():
                    log_queue.put(messages.exiting_by_dialog())
                    stop_event.set()
                    # メニューから終了を選択した場合と同じ動作をさせる
                    # メインスレッドのiconオブジェクトに直接アクセスはできないため、
                    # stop_eventをセットするに留めるのが安全
                    break # 監視ループを抜ける

            # ターゲット情報をリセット
            target_app = target_hwnd = target_filepath = None; last_status_message = ""

        # 4. メイン処理
        if target_app:
            try:
                # --- ▼ tryブロックの範囲はここまで ▼ ---
                vbe_visible = target_app.VBE.MainWindow.Visible
                if vbe_visible:
                    current_modified_time = os.path.getmtime(target_filepath)
                    if current_modified_time > last_modified_time:
                        log_queue.put(messages.formatting_detected(os.path.basename(target_filepath)))
                        last_modified_time = current_modified_time
                        
                        vbe = target_app.VBE
                        for component in vbe.ActiveVBProject.VBComponents:
                            if component.CodeModule.CountOfLines > 0:
                                code_pane = component.CodeModule.CodePane
                                
                                # 1. 現在の位置情報を記憶
                                try:
                                    original_line = code_pane.GetSelection(1, 1, 1, 80)[0]
                                    top_line = code_pane.TopLine
                                except com_error:
                                    original_line = 1; top_line = 1

                                # 2. フォーマット処理
                                original_code = component.CodeModule.Lines(1, component.CodeModule.CountOfLines)
                                formatted_code = formatter.format_code(original_code)
                                
                                if original_code != formatted_code:
                                    component.CodeModule.DeleteLines(1, component.CodeModule.CountOfLines)
                                    component.CodeModule.AddFromString(formatted_code)

                                    # 3. 位置情報を復元
                                    try:
                                        code_pane.TopLine = top_line
                                        code_pane.SetSelection(original_line, 1, original_line, 1)
                                    except com_error:
                                        pass
                        
                        log_queue.put(messages.formatting_complete())
                        last_status_message = ""

                    current_status_message = messages.monitoring_vbe(os.path.basename(target_filepath))
                else:
                    current_status_message = messages.waiting_for_vbe(os.path.basename(target_filepath))

            except com_error: # --- [修正] Exception ではなく com_error を捕捉 ---
                # VBEとの通信でエラーが発生した場合（Excelが強制終了されたなど）
                log_queue.put(messages.connection_lost(os.path.basename(target_filepath)))
                target_app = target_hwnd = target_filepath = None
                last_status_message = ""
        else: # 5. 待機処理
            #print("[DEBUG] No target. Checking window visibility.") # ★追加
            is_visible = is_any_excel_window_visible()
            #print(f"[DEBUG] is_any_excel_window_visible: {is_visible}") # ★追加
            if is_visible: current_status_message = messages.searching_for_book()
            else: current_status_message = messages.waiting_for_excel()

        # (以下、ループの残りは変更なし)
        if current_status_message and current_status_message != last_status_message:
            log_queue.put(current_status_message); last_status_message = current_status_message
        is_visible_now = is_any_excel_window_visible()
        #print(f"[DEBUG] Exit Check: was_visible={was_window_visible}, is_now_visible={is_visible_now}")
        if was_window_visible and not is_visible_now:
            if ask_to_exit():
                log_queue.put(messages.exiting_by_dialog()); stop_event.set(); icon.menu.items[0](icon); break
        was_window_visible = is_visible_now
        time.sleep(1) # ★デバッグ中は sleep を 2 or 3 に伸ばすと追いやすいです

    #print("[DEBUG] monitoring_loop finished") # ★追加
    pythoncom.CoUninitialize()
    log_queue.put(messages.monitoring_stopped())'''