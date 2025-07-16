# SPDX-License-Identifier: MIT
# Copyright (c) 2025 norisan
# active_vba_formatter.py
# v1.0.0 公開
# v1.0.1 二重起動チェック機能
# v1.0.2 公開開始
# v1.0.3 エクセル監視機能を堅牢に。フォーマットした後に1行目に戻るのを修正
# v1.0.4 複数Excelインスタンスの安定した監視に修正
# v1.0.5 ログ出力機能の導入
# v1.0.6 マルチプロセス時のログファイル競合を解消
# v1.0.7 exe化対応および終了処理の堅牢化
# v1.0.8 UIメッセージの改善と終了確認の削除 (リリース最終版)
# ===================================================================================
#
# Version: 1.0.8
#
# 概要:
#   バックグラウンドで常駐し、アクティブなExcelブックのVBAコードを
#   ファイル保存時に自動でインデント整形するツール。
#   OSのUI言語に応じてメッセージを日本語/英語で表示し、タスクトレイから操作可能。
#   動作状況は実行ファイルと同じディレクトリのログファイルに記録される。
#
# 依存ライブラリ:
#   pywin32, pystray, Pillow, psutil
#
# ===================================================================================

import win32com.client
import win32gui
import win32process
import psutil
import os
import time
import pythoncom
import difflib
import sys
import subprocess
import threading
import logging
from logging.handlers import RotatingFileHandler
from pystray import MenuItem as item
import pystray
from PIL import Image, ImageDraw
import tkinter as tk
from tkinter import messagebox
import ctypes
import pywintypes
import winerror
import win32event

# ===================================================================================
# 0. グローバル設定
# ===================================================================================


def func_get_base_dir():
    """
    実行環境（スクリプト or PyInstaller製exe）に応じて基底ディレクトリを返す。
    exeとして実行された場合、sys.executableはexe自身のパスを指す。
    これにより、exeと同じ場所にあるアイコンやログファイルを正しく参照できる。
    """
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))


def func_get_resource_path(relative_path):
    """
    リソースへのパスを解決する。exeにバンドルされたリソースへのパスを正しく取得する。
    PyInstallerで作成されたexeは、リソースを一時フォルダ(_MEIPASS)に展開する。
    """
    if getattr(sys, "frozen", False):
        # exeとして実行されている場合
        base_path = sys._MEIPASS
    else:
        # スクリプトとして実行されている場合
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)


# --- 定数 ---
CHECK_INTERVAL_SECONDS = 2
INDENT_STRING = "    "
BASE_DIR = func_get_base_dir()  # 永続データ（ログファイル）用の基底パス
ICON_FILE_NAME = "active_vba_formatter.ico"
ICON_FILE_PATH = func_get_resource_path(
    ICON_FILE_NAME
)  # バンドルされたリソースのパス解決
LOG_FILE_PATH = os.path.join(BASE_DIR, "active_vba_formatter.log")

# --- グローバルロガー ---
logger = logging.getLogger(__name__)


def func_setup_logging(log_to_file: bool):
    """
    アプリケーションのログ設定。
    log_to_fileフラグにより、ファイルへの出力を制御する。
    これにより、メインプロセスとサブプロセスのログファイルへの書き込み競合を防ぐ。
    """
    logger.setLevel(logging.INFO)
    if logger.hasHandlers():
        logger.handlers.clear()

    log_format = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")

    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(log_format)
    logger.addHandler(console_handler)

    if log_to_file:
        try:
            # ログファイルが5MBを超えたらローテーション（3世代まで保持）
            file_handler = RotatingFileHandler(
                LOG_FILE_PATH, maxBytes=5 * 1024 * 1024, backupCount=3, encoding="utf-8"
            )
            file_handler.setFormatter(log_format)
            logger.addHandler(file_handler)
        except Exception as e:
            logger.error(f"ログファイルハンドラの設定に失敗しました: {e}")


# ===================================================================================
# 1. 多言語メッセージ管理
# ===================================================================================
def func_is_japanese_os() -> bool:
    """OSのUI言語が日本語か判定する。"""
    try:
        return ctypes.windll.kernel32.GetUserDefaultUILanguage() == 1041
    except Exception:
        return False


class Messages:
    """UIメッセージをOS言語に応じて管理するクラス。"""

    def __init__(self):
        self.is_jp = func_is_japanese_os()

    def app_name(self):
        return "VBAフォーマッター" if self.is_jp else "VBA Formatter"

    def menu_quit(self):
        return "終了" if self.is_jp else "Exit"

    def app_is_running(self):
        msg = "{} は既に起動しています。"
        if not self.is_jp:
            msg = "{} is already running."
        return msg.format(self.app_name())

    def startup_error_title(self):
        return "起動エラー" if self.is_jp else "Startup Error"

    def startup_check_error(self, e):
        msg = "二重起動チェック中にエラーが発生しました: {}"
        if not self.is_jp:
            msg = "Error during startup check: {}"
        return msg.format(e)

    def monitoring_started(self):
        return "監視を開始しました。" if self.is_jp else "Monitoring started."

    def watcher_thread_stopped(self):
        return (
            "監視スレッドを終了しました。"
            if self.is_jp
            else "Watcher thread has been stopped."
        )

    def exiting_by_menu(self):
        return (
            "終了メニューがクリックされました。" if self.is_jp else "Exit menu clicked."
        )

    def auto_shutdown_countdown(self):
        if self.is_jp:
            return "Excelウィンドウが全て閉じられました。3秒後に自動終了します..."
        return "All Excel windows have been closed. Shutting down in 3 seconds..."

    def auto_shutdown_now(self):
        return "自動終了します。" if self.is_jp else "Auto-shutting down."

    def all_excel_closed_message(self):
        if self.is_jp:
            return (
                "全てのExcelウィンドウが閉じられたため、アプリケーションを終了します。"
            )
        return "All Excel windows have been closed, so the application will now exit."

    def target_switched(self, f):
        msg = "監視対象を切り替えました: {}"
        if not self.is_jp:
            msg = "Switched target to: {}"
        return msg.format(f)

    def stopped_monitoring_reason(self):
        if self.is_jp:
            return "監視を停止しました (Excelがアクティブではありません)。"
        return "Stopped monitoring (Excel is not active)."

    def file_change_detected(self, f):
        msg = "ファイルの変更を検知: {}"
        if not self.is_jp:
            msg = "File change detected: {}"
        return msg.format(f)

    def launching_formatter(self):
        return "フォーマッタを起動します..." if self.is_jp else "Launching formatter..."

    def formatting_complete(self):
        return (
            "フォーマット処理が完了しました。" if self.is_jp else "Formatting complete."
        )

    def unexpected_error(self, e):
        return f"予期せぬエラー: {e}" if self.is_jp else f"Unexpected error: {e}"

    def formatter_starting(self, f):
        return f"処理開始: {f}" if self.is_jp else f"Processing started: {f}"

    def formatter_component(self, f):
        return f"{f} をフォーマットしました。" if self.is_jp else f"Formatted {f}."

    def formatter_complete_log(self):
        return "処理完了" if self.is_jp else "Processing complete."

    def formatter_error(self):
        return "エラーが発生しました" if self.is_jp else "An error occurred."

    def icon_not_found(self):
        if self.is_jp:
            return f"{os.path.basename(ICON_FILE_PATH)} が見つかりません。ダミー画像を生成します。"
        return (
            f"{os.path.basename(ICON_FILE_PATH)} not found. "
            "Generating a dummy image."
        )

    def notification_message(self):
        if self.is_jp:
            return "起動しました。VBAコードの自動整形を開始します。"
        return "Started. Now monitoring VBA code for auto-formatting."


# ===================================================================================
# 2. VBAコード整形クラス
# ===================================================================================
class VbaFormatter:
    """VBAコードのインデントを整形するロジックを持つクラス。"""

    def __init__(self, indent_char: str = INDENT_STRING):
        self.indent_char = indent_char
        self.INDENT_KEYWORDS = (
            "if",
            "for",
            "do",
            "with",
            "sub",
            "public sub",
            "private sub",
            "function",
            "public function",
            "private function",
            "property",
            "public property",
            "private property",
            "select case",
            "type",
        )
        self.DEDENT_KEYWORDS = (
            "end if",
            "next",
            "loop",
            "end with",
            "end sub",
            "end function",
            "end property",
            "end select",
            "end type",
        )
        self.MID_BLOCK_KEYWORDS = ("else", "elseif", "else if")

    def _func_get_judgement_line(self, code_line: str) -> str:
        """文字列リテラルとコメントを除外した、インデント判断用の行を返す。"""
        clean_line, in_string = "", False
        for char in code_line:
            if char == '"':
                in_string = not in_string
                continue
            if char == "'" and not in_string:
                break
            if not in_string:
                clean_line += char
        return clean_line.strip()

    def func_format_code(self, code_string: str) -> str:
        """与えられたVBAコード文字列を整形して返す。"""
        lines, formatted_lines, current_indent_level, block_stack = (
            code_string.splitlines(),
            [],
            0,
            [],
        )
        for line in lines:
            stripped_line = line.strip()
            if not stripped_line:
                if formatted_lines and formatted_lines[-1] != "":
                    formatted_lines.append("")
                continue
            judgement_line = self._func_get_judgement_line(
                stripped_line.replace("_", "")
            ).lower()
            judgement_parts = judgement_line.split()
            first_word = judgement_parts[0] if judgement_parts else ""
            first_two_words = (
                " ".join(judgement_parts[:2]) if len(judgement_parts) > 1 else ""
            )

            # キーワード判定
            is_start_block = (
                first_two_words in self.INDENT_KEYWORDS
                or first_word in self.INDENT_KEYWORDS
            )
            is_end_block = (
                first_two_words in self.DEDENT_KEYWORDS
                or first_word in self.DEDENT_KEYWORDS
            )
            is_mid_block = (
                first_two_words in self.MID_BLOCK_KEYWORDS
                or first_word in self.MID_BLOCK_KEYWORDS
            )
            is_case_statement = first_word == "case" or first_two_words == "case else"
            is_select_case = first_two_words == "select case"
            is_end_select = first_two_words == "end select"

            # インデントレベルの調整（デデントを先に処理）
            if is_end_select:
                current_indent_level = max(0, current_indent_level - 2)
            elif is_case_statement:
                if block_stack and block_stack[-1] == "in_case":
                    current_indent_level = max(0, current_indent_level - 1)
            elif is_mid_block:
                current_indent_level = max(0, current_indent_level - 1)
            elif is_end_block:
                current_indent_level = max(0, current_indent_level - 1)

            if is_end_select and block_stack:
                block_stack.pop()
            elif is_end_block and block_stack:
                block_stack.pop()

            # 整形後の行を追加
            formatted_lines.append(
                self.indent_char * current_indent_level + stripped_line
            )

            # インデントレベルの調整（インデントを後に処理）
            is_single_line_if = False
            if first_word == "if" and "then" in judgement_line:
                then_pos = judgement_line.find("then")
                rest_of_line = judgement_line[then_pos + 4 :].strip()
                if rest_of_line and not rest_of_line.startswith("'"):
                    is_single_line_if = True

            if is_select_case:
                current_indent_level += 1
                block_stack.append("select")
            elif is_case_statement:
                current_indent_level += 1
                if block_stack and block_stack[-1] == "select":
                    block_stack[-1] = "in_case"
            elif (is_start_block and not is_single_line_if) or is_mid_block:
                current_indent_level += 1
                if is_start_block and not is_single_line_if:
                    block_stack.append("other")

        # 余分な空行を削除
        final_lines = []
        for i, line in enumerate(formatted_lines):
            if line == "" and (i == 0 or final_lines[-1] == ""):
                continue
            final_lines.append(line)
        return "\n".join(final_lines)


# ===================================================================================
# 3. ヘルパー関数群
# ===================================================================================
VBA_FORMATTER_INSTANCE = VbaFormatter()


def func_format_vba_code(code_string: str) -> str:
    """VBA整形インスタンスを介してコードをフォーマットする。"""
    return VBA_FORMATTER_INSTANCE.func_format_code(code_string)


def func_create_dummy_image():
    """アイコンファイルが見つからない場合にダミーの画像を生成する。"""
    width, height, color1, color2 = 64, 64, "black", "white"
    image = Image.new("RGB", (width, height), color1)
    dc = ImageDraw.Draw(image)
    dc.rectangle((width // 2, 0, width, height // 2), fill=color2)
    dc.rectangle((0, height // 2, width // 2, height), fill=color2)
    return image


def func_find_visible_excel_windows():
    """表示されている全てのExcelウィンドウのハンドルをリストで返す。"""
    visible_excel_windows = []

    def _func_enum_windows_callback(hwnd, _):
        if win32gui.IsWindowVisible(hwnd) and win32gui.GetClassName(hwnd) == "XLMAIN":
            visible_excel_windows.append(hwnd)

    win32gui.EnumWindows(_func_enum_windows_callback, None)
    return visible_excel_windows


def func_show_windows_messagebox(title, message, style):
    """
    Tkinterに依存しない、Windows APIを直接呼び出すメッセージボックス。
    exe化されたアプリがGUIメインループ開始前にメッセージを出す際の安定性を確保する。
    style: 0 = OK, 16 = Stop icon, 48 = Warning icon
    """
    return ctypes.windll.user32.MessageBoxW(0, message, title, style)


# ===================================================================================
# 4. フォーマット実行役 (サブプロセス側)
# ===================================================================================
def func_apply_formatting_to_active_excel():
    """
    サブプロセスとして起動され、アクティブなExcelインスタンスに接続し、
    VBAコードのフォーマットを実行する。
    """
    messages = Messages()
    try:
        pythoncom.CoInitialize()
        excel_app = win32com.client.GetActiveObject("Excel.Application")
        workbook = excel_app.ActiveWorkbook
        if not workbook or not workbook.Name:
            return

        logger.info(f"--- [Formatter] {messages.formatter_starting(workbook.Name)} ---")
        vb_project = workbook.VBProject
        for component in vb_project.VBComponents:
            module = component.CodeModule
            if module.CountOfLines == 0:
                continue

            original_code = module.Lines(1, module.CountOfLines)
            formatted_code = func_format_vba_code(original_code)

            if original_code.splitlines() == formatted_code.splitlines():
                continue

            # 差分を検出し、必要な部分だけを置換
            matcher = difflib.SequenceMatcher(
                None, original_code.splitlines(), formatted_code.splitlines()
            )
            for tag, i1, i2, j1, j2 in reversed(matcher.get_opcodes()):
                if tag == "equal":
                    continue

                start_line = i1 + 1
                if tag == "replace":
                    module.DeleteLines(start_line, i2 - i1)
                    module.InsertLines(
                        start_line, "\n".join(formatted_code.splitlines()[j1:j2])
                    )
                elif tag == "delete":
                    module.DeleteLines(start_line, i2 - i1)
                elif tag == "insert":
                    module.InsertLines(
                        start_line, "\n".join(formatted_code.splitlines()[j1:j2])
                    )

            logger.info(f"  -> {messages.formatter_component(component.Name)}")
        logger.info(f"--- [Formatter] {messages.formatter_complete_log()} ---")
    except Exception:
        logger.exception(f"---!!! [Formatter] {messages.formatter_error()} !!!---")
    finally:
        pythoncom.CoUninitialize()


# ===================================================================================
# 5. 監視役アプリケーションクラス
# ===================================================================================
class WatcherApp:
    """タスクトレイ常駐、ファイル監視、サブプロセス起動を管理するメインクラス。"""

    def __init__(self, messages_instance):
        self.messages = messages_instance
        self.stop_event = threading.Event()
        self.tray_icon = None
        self.root = tk.Tk()
        self.root.withdraw()
        self.root.wm_attributes("-topmost", 1)
        self.watcher_thread = None  # 監視スレッドの参照を保持

    def func_run_watcher_thread(self):
        """アクティブウィンドウとファイル変更を監視するバックグラウンドスレッド。"""
        pythoncom.CoInitialize()
        logger.info(f"[Watcher] {self.messages.monitoring_started()}")
        monitored_file, last_mod_time, excel_closed_time, has_excel_run = (
            None,
            0,
            None,
            False,
        )

        while not self.stop_event.is_set():
            try:
                if self.stop_event.wait(CHECK_INTERVAL_SECONDS):
                    break

                if not func_find_visible_excel_windows():
                    if not has_excel_run:
                        continue
                    if excel_closed_time is None:
                        excel_closed_time = time.time()
                        logger.info(
                            f"[Watcher] {self.messages.auto_shutdown_countdown()}"
                        )
                    if time.time() - excel_closed_time > 3:
                        logger.info(f"[Watcher] {self.messages.auto_shutdown_now()}")
                        if self.tray_icon:
                            self.tray_icon.notify(
                                self.messages.all_excel_closed_message(),
                                self.messages.app_name(),
                            )
                        time.sleep(4)
                        if self.tray_icon:
                            self.tray_icon.stop()
                        break
                    continue
                else:
                    has_excel_run = True
                    excel_closed_time = None

                current_file_path = None
                excel = None
                try:
                    hwnd = win32gui.GetForegroundWindow()
                    if hwnd != 0:
                        _, pid = win32process.GetWindowThreadProcessId(hwnd)
                        if psutil.Process(pid).name().lower() == "excel.exe":
                            excel = win32com.client.GetActiveObject("Excel.Application")
                            if excel.ActiveWorkbook:
                                current_file_path = excel.ActiveWorkbook.FullName
                except (
                    psutil.NoSuchProcess,
                    pythoncom.com_error,
                    AttributeError,
                    pywintypes.error,
                ):
                    pass
                finally:
                    if excel:
                        excel = None

                if not current_file_path:
                    if monitored_file:
                        logger.info(
                            f"[Watcher] {self.messages.stopped_monitoring_reason()}"
                        )
                        monitored_file = None
                    continue

                if current_file_path != monitored_file:
                    monitored_file = current_file_path
                    if os.path.exists(monitored_file):
                        last_mod_time = os.path.getmtime(monitored_file)
                        logger.info(
                            f"[Watcher] {self.messages.target_switched(os.path.basename(monitored_file))}"
                        )

                if monitored_file and os.path.exists(monitored_file):
                    current_mod_time = os.path.getmtime(monitored_file)
                    if current_mod_time != last_mod_time:
                        last_mod_time = current_mod_time
                        logger.info(
                            f"[Watcher] {self.messages.file_change_detected(os.path.basename(monitored_file))}"
                        )

                        # exe化された環境とスクリプト実行環境でコマンドを分岐させる
                        if getattr(sys, "frozen", False):
                            # exeとして実行されている場合: ["VBA_Formatter.exe", "--format-now"]
                            cmd = [sys.executable, "--format-now"]
                        else:
                            # スクリプトとして実行されている場合: ["python.exe", "active_vba_formatter.py", "--format-now"]
                            cmd = [sys.executable, __file__, "--format-now"]

                        logger.info(f"[Watcher] {self.messages.launching_formatter()}")
                        subprocess.run(
                            cmd, check=True, creationflags=subprocess.CREATE_NO_WINDOW
                        )

                        logger.info(f"[Watcher] {self.messages.formatting_complete()}")
                        time.sleep(0.5)
                        if os.path.exists(monitored_file):
                            last_mod_time = os.path.getmtime(monitored_file)

            except Exception as e:
                logger.exception(f"[Watcher] {self.messages.unexpected_error(e)}")
                monitored_file = None
                time.sleep(5)

        pythoncom.CoUninitialize()
        logger.info(f"[Watcher] {self.messages.watcher_thread_stopped()}")

    def func_setup_and_run_tray(self):
        """タスクトレイアイコンを設定し、監視スレッドを開始する。"""
        try:
            image = Image.open(ICON_FILE_PATH)
        except FileNotFoundError:
            logger.warning(self.messages.icon_not_found())
            image = func_create_dummy_image()

        menu = (item(self.messages.menu_quit(), self.func_exit_app),)
        self.tray_icon = pystray.Icon(
            "VBA Formatter", image, self.messages.app_name(), menu
        )

        self.watcher_thread = threading.Thread(
            target=self.func_run_watcher_thread, daemon=True
        )
        self.watcher_thread.start()

        self.tray_icon.run(setup=self.func_show_startup_notification)

    def func_show_startup_notification(self, icon):
        """起動時にバルーン通知を表示する。"""
        icon.visible = True
        icon.notify(self.messages.notification_message(), self.messages.app_name())

    def func_exit_app(self):
        """手動でのアプリケーション終了処理。確認ダイアログは表示しない。"""
        logger.info(f"[Watcher] {self.messages.exiting_by_menu()}")

        self.stop_event.set()

        # 監視スレッドが完全に終了するのを待つことで、クリーンな終了を保証する
        if self.watcher_thread and self.watcher_thread.is_alive():
            self.watcher_thread.join(timeout=CHECK_INTERVAL_SECONDS + 1)

        if self.tray_icon:
            self.tray_icon.stop()
        if self.root:
            self.root.destroy()


# ===================================================================================
# 6. 起動ロジック
# ===================================================================================
if __name__ == "__main__":
    # 実行時引数で「監視役」か「整形役」かを判断
    is_formatter_process = len(sys.argv) > 1 and sys.argv[1] == "--format-now"

    if is_formatter_process:
        # 整形役（サブプロセス）の場合、ログはコンソールにのみ出力
        func_setup_logging(log_to_file=False)
        func_apply_formatting_to_active_excel()
    else:
        # 監視役（メインプロセス）の場合、ログをファイルにも出力
        func_setup_logging(log_to_file=True)

        # ミューテックスを使い、二重起動を防止する
        messages = Messages()
        MUTEX_NAME = "VBAFormatter_Mutex_{C1A9E7D8-1B2C-4F3D-9A8E-5G6H7I8J9K0L}"
        mutex = None
        try:
            logger.info("アプリケーションを起動します。ミューテックスを確認中...")
            mutex = win32event.CreateMutex(None, 1, MUTEX_NAME)
        except pywintypes.error as e:
            if e.winerror == winerror.ERROR_ALREADY_EXISTS:
                logger.warning(
                    "ミューテックスが既に存在するため、二重起動と判断しました。"
                )
                # Tkinterをこの場で初期化し、アイコンを設定してメッセージボックスを表示
                root = tk.Tk()
                root.withdraw()
                try:
                    # バンドルされたアイコンリソースへのパスを使用
                    icon_path = func_get_resource_path(ICON_FILE_NAME)
                    if os.path.exists(icon_path):
                        root.iconbitmap(icon_path)
                except Exception:
                    pass  # アイコン設定に失敗しても続行
                root.wm_attributes("-topmost", 1)
                messagebox.showwarning(
                    messages.app_name(), messages.app_is_running(), parent=root
                )
                root.destroy()
                sys.exit(0)
            else:
                logger.exception(
                    "ミューテックスの作成中に予期せぬエラーが発生しました。"
                )
                root = tk.Tk()
                root.withdraw()
                try:
                    icon_path = func_get_resource_path(ICON_FILE_NAME)
                    if os.path.exists(icon_path):
                        root.iconbitmap(icon_path)
                except Exception:
                    pass
                root.wm_attributes("-topmost", 1)
                error_msg = messages.startup_check_error(e)
                messagebox.showerror(
                    messages.startup_error_title(), error_msg, parent=root
                )
                root.destroy()
                sys.exit(1)

        logger.info("ミューテックスの作成に成功。監視アプリケーションを起動します。")
        app = WatcherApp(messages)
        app.func_setup_and_run_tray()

        if mutex:
            win32event.ReleaseMutex(mutex)
            logger.info("ミューテックスを解放しました。")

        logger.info("[Watcher] プログラムを完全に終了しました。")
