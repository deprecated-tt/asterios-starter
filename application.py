import sys
import os
import json
import time
import psutil
import threading
import queue
import platform
from functools import partial
from PyQt5 import QtWidgets, QtGui, QtCore
from PyQt5.QtWidgets import (
    QFileDialog, QInputDialog, QWidget, QHBoxLayout, QCheckBox, QPushButton, QWidgetAction
)

SETTINGS_FILE = "settings.json"
KEYS_FILE = "saved_keys.txt"
TARGET_PROCESS_NAME = "AsteriosGame.exe"
AUTOSTART_LIMIT = 3
IS_WINDOWS = platform.system().lower() == "windows"

# ---------------- Elevation (UAC) ----------------
if IS_WINDOWS:
    import ctypes
    from ctypes import wintypes

    def _pythonw_path():
        # Пытаемся найти pythonw.exe рядом с текущим интерпретатором (venv/…/pythonw.exe)
        cand1 = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
        cand2 = sys.executable.replace("python.exe", "pythonw.exe")
        for p in (cand1, cand2):
            if os.path.exists(p):
                return p
        return sys.executable  # fallback (будет консоль, но это крайний случай)

    def is_admin():
        try:
            return ctypes.windll.shell32.IsUserAnAdmin()
        except Exception:
            return False

    def relaunch_as_admin():
        # Релончим через pythonw.exe, чтобы не открывалась консоль
        interpreter = _pythonw_path()
        params = " ".join([f'"{a}"' if " " in a else a for a in sys.argv])
        ret = ctypes.windll.shell32.ShellExecuteW(None, "runas", interpreter, params, None, 1)
        return ret > 32

    def ensure_admin():
        if not is_admin():
            if not relaunch_as_admin():
                raise SystemExit("Не удалось запросить права администратора.")
            raise SystemExit(0)

# ---------------- WinAPI helpers ----------------
if IS_WINDOWS:
    import ctypes
    from ctypes import wintypes

    user32 = ctypes.WinDLL('user32', use_last_error=True)
    shell32 = ctypes.WinDLL('shell32', use_last_error=True)

    # DPI awareness
    def _enable_dpi_awareness():
        try:
            SetProcessDpiAwarenessContext = user32.SetProcessDpiAwarenessContext
            SetProcessDpiAwarenessContext.restype = wintypes.BOOL
            if SetProcessDpiAwarenessContext(ctypes.c_void_p(-4)):  # PER_MONITOR_AWARE_V2
                return
        except Exception:
            pass
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(2)  # PER_MONITOR
            return
        except Exception:
            pass
        try:
            user32.SetProcessDPIAware()
        except Exception:
            pass

    WNDENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

    EnumWindows = user32.EnumWindows
    GetWindowThreadProcessId = user32.GetWindowThreadProcessId
    IsWindowVisible = user32.IsWindowVisible
    GetForegroundWindow = user32.GetForegroundWindow
    GetWindowRect = user32.GetWindowRect
    ShowWindow = user32.ShowWindow
    SetWindowPos = user32.SetWindowPos
    ShellExecuteW = shell32.ShellExecuteW
    GetWindowLongW = user32.GetWindowLongW
    GetWindow = user32.GetWindow

    # consts
    SW_RESTORE = 9
    HWND_TOP = wintypes.HWND(0)
    SWP_NOZORDER   = 0x0004
    SWP_NOACTIVATE = 0x0010
    SWP_SHOWWINDOW = 0x0040
    GW_OWNER = 4
    GWL_EXSTYLE = -20
    WS_EX_TOOLWINDOW = 0x00000080  # tool windows

class LaunchTask:
    def __init__(self, name: str, key_value: str):
        self.name = name
        self.key_value = key_value

class TrayApp(QtWidgets.QSystemTrayIcon):
    def __init__(self, icon, parent=None):
        super().__init__(icon, parent)
        self.setToolTip("Key Manager")

        self.file_path = None
        self.exe_path = None
        self.autostart_keys = []
        self.window_positions = {}
        self.saved_key_items = []
        self.checkbox_by_name = {}

        self.task_queue = queue.Queue()
        self.worker_thread = threading.Thread(target=self.process_queue, daemon=True)
        self.worker_thread.start()

        self.menu = QtWidgets.QMenu(parent)
        self.setContextMenu(self.menu)

        self.keys_placeholder = self.menu.addAction("— No saved keys —")
        self.keys_placeholder.setEnabled(False)

        self.menu.addSeparator()
        self.extract_action = self.menu.addAction("💾 Extract Key from File")

        self.settings_menu = QtWidgets.QMenu("⚙️ Settings", self.menu)
        self.config_action = self.settings_menu.addAction("📂 Configure File")
        self.exe_config_action = self.settings_menu.addAction("⚙️ Configure EXE")
        self.menu.addMenu(self.settings_menu)

        self.menu.addSeparator()
        self.quit_action = self.menu.addAction("❌ Quit")

        self.config_action.triggered.connect(self.select_file)
        self.exe_config_action.triggered.connect(self.select_exe_file)
        self.extract_action.triggered.connect(self.extract_key)
        self.quit_action.triggered.connect(QtWidgets.qApp.quit)

        self.load_settings()
        self.refresh_saved_keys_menu()
        self.enqueue_autostart_keys()

    # ---------- Settings ----------
    def load_settings(self):
        if os.path.exists(SETTINGS_FILE):
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                self.file_path = data.get("file_path")
                self.exe_path = data.get("exe_path")
                self.autostart_keys = data.get("autostart_keys", []) or []
                self.window_positions = data.get("window_positions", {}) or {}

    def save_settings(self):
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump({
                "file_path": self.file_path,
                "exe_path": self.exe_path,
                "autostart_keys": self.autostart_keys,
                "window_positions": self.window_positions
            }, f, ensure_ascii=False, indent=2)

    # ---------- File pickers ----------
    def select_file(self):
        file_path, _ = QFileDialog.getOpenFileName(None, "Select Config File")
        if file_path:
            self.file_path = file_path
            self.save_settings()

    def select_exe_file(self):
        exe_path, _ = QFileDialog.getOpenFileName(None, "Select EXE File", filter="Executable (*.exe)")
        if exe_path:
            self.exe_path = exe_path
            self.save_settings()

    # ---------- Keys I/O ----------
    def load_keys_map(self):
        keys = {}
        if not os.path.exists(KEYS_FILE):
            return keys
        with open(KEYS_FILE, "r") as f:
            for line in f:
                line = line.strip()
                if not line or "=" not in line:
                    continue
                name, value = line.split("=", 1)
                keys[name] = value
        return keys

    def extract_key(self):
        if not self.file_path:
            return
        with open(self.file_path, 'r') as f:
            key_line = next((line.strip() for line in f if line.startswith("Key=")), None)
        if not key_line:
            return
        key_value = key_line.split("=", 1)[1]
        name, ok = QInputDialog.getText(None, "Save Key", "Enter name to save this key under:")
        if ok and name:
            with open(KEYS_FILE, "a") as f:
                f.write(f"{name}={key_value}\n")
            self.refresh_saved_keys_menu()

    # ---------- Menu build ----------
    def clear_keys_menu(self):
        for item in self.saved_key_items:
            self.menu.removeAction(item)
        self.saved_key_items.clear()
        self.checkbox_by_name.clear()
        self.keys_placeholder.setVisible(False)

    def make_key_row(self, name, value):
        row_widget = QWidget()
        row_layout = QHBoxLayout(row_widget)
        row_layout.setContentsMargins(4, 0, 4, 0)
        row_layout.setSpacing(6)

        cb = QCheckBox()
        cb.setChecked(name in self.autostart_keys)
        cb.toggled.connect(partial(self.toggle_autostart_ui_guard, name))
        self.checkbox_by_name[name] = cb

        btn = QPushButton(name)
        btn.clicked.connect(partial(self.enqueue_task, name, value))
        btn.setFlat(True)
        btn.setStyleSheet(
            "QPushButton {border:none; text-align:left; padding:6px;} "
            "QPushButton:hover {background:palette(Highlight); color:palette(HighlightedText);}"
        )
        btn.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Preferred)

        save_btn = QPushButton("💾")
        save_btn.setToolTip("Save window position for this key")
        save_btn.setFlat(True)
        save_btn.clicked.connect(partial(self.save_position_for_key, name))

        row_layout.addWidget(cb)
        row_layout.addWidget(btn, 1)
        row_layout.addWidget(save_btn)

        wa = QWidgetAction(self.menu)
        wa.setDefaultWidget(row_widget)
        return wa

    def refresh_saved_keys_menu(self):
        self.clear_keys_menu()
        keys_map = self.load_keys_map()
        if not keys_map:
            self.keys_placeholder.setVisible(True)
            return
        insert_before = self.menu.actions()[0] if self.menu.actions() else None
        for name, value in keys_map.items():
            wa = self.make_key_row(name, value)
            if insert_before:
                self.menu.insertAction(insert_before, wa)
            else:
                self.menu.addAction(wa)
            self.saved_key_items.append(wa)

    # ---------- Autostart ----------
    def toggle_autostart_ui_guard(self, name, checked):
        if checked:
            if name not in self.autostart_keys and len(self.autostart_keys) < AUTOSTART_LIMIT:
                self.autostart_keys.append(name)
        else:
            if name in self.autostart_keys:
                self.autostart_keys.remove(name)
        self.save_settings()

    def enqueue_autostart_keys(self):
        keys_map = self.load_keys_map()
        for name in self.autostart_keys[:AUTOSTART_LIMIT]:
            if name in keys_map:
                self.enqueue_task(name, keys_map[name])

    # ---------- Queue / launch ----------
    def enqueue_task(self, name, value):
        self.task_queue.put(LaunchTask(name, value))

    def process_queue(self):
        while True:
            task = self.task_queue.get()
            if task:
                self.perform_launch(task.name, task.key_value)
            self.task_queue.task_done()

    def perform_launch(self, name, value):
        if not self.file_path or not self.exe_path or not os.path.exists(self.exe_path):
            return

        # Подставляем Key=...
        with open(self.file_path, "r") as f:
            lines = f.readlines()
        updated = False
        for i, line in enumerate(lines):
            if line.startswith("Key="):
                lines[i] = f"Key={value}\n"
                updated = True
                break
        if not updated:
            lines.append(f"Key={value}\n")
        with open(self.file_path, "w") as f:
            f.writelines(lines)

        existing_pids = self.get_existing_game_pids()
        self.start_client_process(self.exe_path, "/autoplay")

        new_pid = self.wait_for_new_game_process(existing_pids)
        if new_pid and IS_WINDOWS:
            self.apply_saved_window_position(name, new_pid)

    def start_client_process(self, exe_path: str, params: str = ""):
        if not IS_WINDOWS:
            import subprocess
            subprocess.Popen([exe_path] + ([params] if params else []), shell=False)
            return
        workdir = os.path.dirname(exe_path) or None
        if is_admin():
            import subprocess
            subprocess.Popen([exe_path] + ([params] if params else []), cwd=workdir, shell=False)
        else:
            # если трей вдруг не elevated (не должен, но на всякий случай)
            ret = ShellExecuteW(None, "runas", exe_path, params, workdir, 1)
            if ret <= 32:
                raise OSError(f"ShellExecuteW failed with code {ret}")

    def get_existing_game_pids(self):
        return {p.info['pid'] for p in psutil.process_iter(['pid', 'name'])
                if p.info['name'] and p.info['name'].lower() == TARGET_PROCESS_NAME.lower()}

    def wait_for_new_game_process(self, existing_pids, timeout=30):
        deadline = time.time() + timeout
        newest_pid, newest_ct = None, -1.0
        while time.time() < deadline:
            for p in psutil.process_iter(['pid', 'name', 'create_time']):
                if p.info['name'] and p.info['name'].lower() == TARGET_PROCESS_NAME.lower():
                    if p.info['pid'] in existing_pids:
                        continue
                    ct = p.info.get('create_time') or 0.0
                    if ct > newest_ct:
                        newest_ct, newest_pid = ct, p.info['pid']
            if newest_pid:
                time.sleep(2.0)  # дать окну успеть появиться
                return newest_pid
            time.sleep(0.3)
        return None

    # ---------- Window position ----------
    def save_position_for_key(self, name):
        """Сохраняем позицию ИМЕННО игрового окна (игнорируя загрузчик)."""
        if not IS_WINDOWS:
            return

        # если игра уже запущена — берём её «игровое» окно; иначе — foreground
        pid = None
        newest_ct = -1.0
        for p in psutil.process_iter(['pid', 'name', 'create_time']):
            if p.info['name'] and p.info['name'].lower() == TARGET_PROCESS_NAME.lower():
                ct = p.info.get('create_time') or 0.0
                if ct > newest_ct:
                    newest_ct, pid = ct, p.info['pid']

        hwnd = self._pick_gameplay_window(pid) if pid else GetForegroundWindow()
        rect = self._get_window_rect(hwnd)
        if not rect:
            return

        x, y, w, h = rect

        self.window_positions[name] = {"x": x, "y": y, "width": w, "height": h}
        self.save_settings()

    def apply_saved_window_position(self, name, pid):
        """Ждём/находим игровое окно (не загрузчик) и двигаем его несколько раз (до 7 сек)."""
        pos = self.window_positions.get(name)
        if not pos or not IS_WINDOWS:
            return
        target_x, target_y, target_w, target_h = map(int, (pos["x"], pos["y"], pos["width"], pos["height"]))

        deadline = time.time() + 20.0
        last_hwnd = None
        while time.time() < deadline:
            hwnd = self._pick_gameplay_window(pid)
            if hwnd and int(hwnd) != 0:
                last_hwnd = hwnd
                ShowWindow(hwnd, SW_RESTORE)
                ok = SetWindowPos(hwnd, HWND_TOP, target_x, target_y, target_w, target_h,
                                  SWP_NOZORDER | SWP_NOACTIVATE | SWP_SHOWWINDOW)
                if not ok:
                    # это не откроет консоль, т.к. процесс pythonw.exe — но полезно оставить на случай лога в файл
                    err = ctypes.get_last_error()
                    print(f"SetWindowPos failed, GetLastError={err}")
                cur = self._get_window_rect(hwnd)
                if cur and abs(cur[0]-target_x) <= 2 and abs(cur[1]-target_y) <= 2:
                    return
            time.sleep(0.5)

        if last_hwnd:
            ShowWindow(last_hwnd, SW_RESTORE)
            SetWindowPos(last_hwnd, HWND_TOP, target_x, target_y, target_w, target_h,
                         SWP_NOZORDER | SWP_NOACTIVATE | SWP_SHOWWINDOW)

    # ---- helpers ----
    def _get_window_rect(self, hwnd):
        if not hwnd:
            return None
        rect = wintypes.RECT()
        if not GetWindowRect(hwnd, ctypes.byref(rect)):
            return None
        return (rect.left, rect.top, rect.right-rect.left, rect.bottom-rect.top)

    def _is_tool_or_owned(self, hwnd):
        owner = GetWindow(hwnd, GW_OWNER)
        if owner:
            return True
        exs = GetWindowLongW(hwnd, GWL_EXSTYLE)
        return bool(exs & WS_EX_TOOLWINDOW)

    def _pick_gameplay_window(self, pid, min_w=0, min_h=0):
        """Выбираем видимое top-level окно процесса с максимальной площадью, без owner/tool, и не меньше порога."""
        if not IS_WINDOWS or not pid:
            return None
        best_hwnd, best_area = None, -1

        @WNDENUMPROC
        def enum_proc(hwnd, lParam):
            nonlocal best_hwnd, best_area
            if not IsWindowVisible(hwnd):
                return True
            proc_pid = wintypes.DWORD(0)
            GetWindowThreadProcessId(hwnd, ctypes.byref(proc_pid))
            if int(proc_pid.value) != int(pid):
                return True
            if self._is_tool_or_owned(hwnd):
                return True
            rect = wintypes.RECT()
            if not GetWindowRect(hwnd, ctypes.byref(rect)):
                return True
            w, h = rect.right - rect.left, rect.bottom - rect.top
            if w < min_w or h < min_h:
                return True
            area = w * h
            if area > best_area:
                best_area = area
                best_hwnd = hwnd
            return True

        EnumWindows(enum_proc, 0)
        return best_hwnd


# ---------- App bootstrap ----------
def resource_path(relative_path):
    base_path = getattr(sys, '_MEIPASS', os.path.abspath("."))
    return os.path.join(base_path, relative_path)

def main():
    if IS_WINDOWS:
        ensure_admin()          # теперь релонч через pythonw.exe -> без чёрного окна
        _enable_dpi_awareness()
    app = QtWidgets.QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)
    icon_path = resource_path("icon.png")
    icon = QtGui.QIcon(icon_path) if os.path.exists(icon_path) else QtGui.QIcon()
    tray = TrayApp(icon)
    tray.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
