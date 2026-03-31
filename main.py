import sys
import os
import json
import winreg
import subprocess
import webbrowser
import ctypes
from datetime import datetime
from PyQt6.QtCore import Qt, QTimer, QUrl
from PyQt6.QtNetwork import QLocalServer, QLocalSocket
from PyQt6.QtGui import QIcon, QAction
from PyQt6.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QSystemTrayIcon, QMenu, QTableWidgetItem, QHeaderView)
from PyQt6.QtMultimedia import QMediaPlayer, QAudioOutput
from qfluentwidgets import (FluentWindow, NavigationItemPosition, MessageBox, 
                            SubtitleLabel, BodyLabel, CaptionLabel, LineEdit, PushButton, ComboBox, 
                            SwitchButton, setTheme, Theme, TableWidget, 
                            TimePicker, TextEdit, InfoBar, MessageBoxBase, FluentIcon as FIF,
                            HorizontalSeparator)

# --- 경로 설정 함수 ---
def get_base_path():
    """ 프로그램의 실행 파일 또는 스크립트가 위치한 디렉토리 경로를 반환합니다. """
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

def get_resource_path(relative_path):
    """ PyInstaller 환경(리소스 번들)에서도 파일 경로를 올바르게 반환합니다. """
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(get_base_path(), relative_path)

# --- 설정 및 상수 ---
BASE_PATH = get_base_path()
CONFIG_FILE = os.path.join(BASE_PATH, "config.json")
SOUND_FILE = get_resource_path(os.path.join("sounds", "Alarm.mp3"))
APP_NAME = "AnuAutoScheduler"
SERVER_NAME = "AnuAutoScheduler_Unique_Server"
VERSION = "1.0.1"

class ConfigManager:
    @staticmethod
    def load():
        default_config = {"run_at_startup": False, "theme": "Light", "browser": "Chrome", "schedules": []}
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    # 누락된 기본 키 보강
                    for key, value in default_config.items():
                        if key not in data:
                            data[key] = value
                    return data
            except (json.JSONDecodeError, IOError):
                return default_config
        return default_config

    @staticmethod
    def save(data):
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=4, ensure_ascii=False)
        except IOError as e:
            print(f"Failed to save config: {e}")

class RegistryUtils:
    @staticmethod
    def set_startup(enable):
        key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
        
        # 실행 파일 경로 확보 및 큰따옴표 처리
        if getattr(sys, 'frozen', False):
            exe_path = f'"{sys.executable}"'
        else:
            python_exe = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
            if not os.path.exists(python_exe):
                python_exe = sys.executable
            script_path = os.path.abspath(sys.argv[0])
            exe_path = f'"{python_exe}" "{script_path}"'

        try:
            # KEY_WRITE와 KEY_WOW64_64KEY를 사용하여 64비트 시스템에서도 안정적으로 쓰기 권한 확보
            access = winreg.KEY_WRITE
            if sys.maxsize > 2**32: # 64비트 환경 체크
                access |= winreg.KEY_WOW64_64KEY
            
            key = winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, key_path, 0, access)
            if enable:
                winreg.SetValueEx(key, APP_NAME, 0, winreg.REG_SZ, exe_path)
            else:
                try:
                    winreg.DeleteValue(key, APP_NAME)
                except FileNotFoundError:
                    pass
            winreg.CloseKey(key)
            return True
        except Exception as e:
            print(f"Startup registry error: {e}")
            return False

    @staticmethod
    def get_browser_path(browser="Chrome"):
        path_map = {
            "Chrome": r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe",
            "Edge": r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe"
        }
        reg_path = path_map.get(browser, path_map["Chrome"])
        for flags in [winreg.KEY_READ | winreg.KEY_WOW64_64KEY, winreg.KEY_READ | winreg.KEY_WOW64_32KEY]:
            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path, 0, flags) as key:
                    return winreg.QueryValue(key, None)
            except WindowsError:
                continue
        return None

class ShortcutUtils:
    @staticmethod
    def create_desktop_shortcut():
        """ 바탕화면에 바로가기 아이콘을 생성합니다. """
        try:
            desktop = os.path.join(os.environ['USERPROFILE'], 'Desktop')
            shortcut_path = os.path.join(desktop, f"{APP_NAME}.lnk")
            
            if getattr(sys, 'frozen', False):
                target_path = sys.executable
                arguments = ""
                icon_location = sys.executable
                working_dir = os.path.dirname(sys.executable)
            else:
                python_exe = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
                if not os.path.exists(python_exe):
                    python_exe = sys.executable
                target_path = python_exe
                arguments = os.path.abspath(sys.argv[0])
                icon_location = python_exe
                working_dir = os.path.dirname(arguments)

            # PowerShell을 사용하여 바로가기 생성
            ps_command = (
                f'$WshShell = New-Object -ComObject WScript.Shell; '
                f'$Shortcut = $WshShell.CreateShortcut("{shortcut_path}"); '
                f'$Shortcut.TargetPath = "{target_path}"; '
                f'$Shortcut.Arguments = "{arguments}"; '
                f'$Shortcut.WorkingDirectory = "{working_dir}"; '
                f'$Shortcut.IconLocation = "{icon_location}"; '
                f'$Shortcut.Save()'
            )
            
            subprocess.run(["powershell", "-Command", ps_command], capture_output=True, check=True)
            return True
        except Exception as e:
            print(f"Shortcut creation error: {e}")
            return False

# --- UI 컴포넌트들 ---

class AddScheduleDialog(MessageBoxBase):
    """ 새 스케줄을 추가하는 팝업 다이얼로그 """
    def __init__(self, main_window):
        super().__init__(main_window)
        self.main_window = main_window
        self.setObjectName("AddScheduleDialog")
        
        self.titleLabel = SubtitleLabel("새 스케줄 추가", self)
        self.viewLayout.addWidget(self.titleLabel)

        self.time_picker = TimePicker(self)
        self.viewLayout.addWidget(BodyLabel("시간 설정 (시:분):"))
        self.viewLayout.addWidget(self.time_picker)

        self.task_name_input = LineEdit(self)
        self.task_name_input.setPlaceholderText("작업 이름 입력")
        self.viewLayout.addWidget(self.task_name_input)

        self.url_input = LineEdit(self)
        self.url_input.setPlaceholderText("접속할 URL 입력 (https:// 포함)")
        self.viewLayout.addWidget(self.url_input)

        self.script_input = TextEdit(self)
        self.script_input.setPlaceholderText("자동화 스크립트 (추후 지원 예정)")
        self.viewLayout.addWidget(self.script_input)

        self.yesButton.setText("스케줄 저장")
        self.cancelButton.setText("취소")
        
        self.widget.setMinimumWidth(350)
        
        # yesButton 클릭 시 기본 accept() 대신 유효성 검사 실행
        self.yesButton.clicked.disconnect()
        self.yesButton.clicked.connect(self.validate_and_save)

    def validate_and_save(self):
        time = self.time_picker.getTime()
        hour, minute = time.hour(), time.minute()
        task_name = self.task_name_input.text().strip()
        url = self.url_input.text().strip()
        
        if not task_name or not url:
            InfoBar.warning("입력 부족", "작업 이름과 URL을 모두 입력해주세요.", parent=self.main_window)
            return

        if '.' not in url or len(url) < 4:
            InfoBar.warning("URL 오류", "올바른 URL 형식이 아닙니다. (예: https://google.com)", parent=self.main_window)
            return

        duplicate = any(s['hour'] == hour and s['minute'] == minute for s in self.main_window.config['schedules'])
        if duplicate:
            msg = MessageBox("동일 시간 스케줄", 
                             f"해당 시간({hour:02d}:{minute:02d})에 이미 등록된 스케줄이 있습니다.\n추가하시겠습니까?", 
                             self.main_window)
            if not msg.exec():
                return

        new_sched = {
            "hour": hour, "minute": minute,
            "task_name": task_name,
            "url": url,
            "script": self.script_input.toPlainText(),
            "last_run_date": ""
        }
        self.main_window.config['schedules'].append(new_sched)
        ConfigManager.save(self.main_window.config)
        self.main_window.home_widget.update_table()
        InfoBar.success("성공", "스케줄이 성공적으로 추가되었습니다.", parent=self.main_window)
        
        self.accept()  # 성공 시 다이얼로그 닫기

class HomeWidget(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        self.setObjectName("HomeWidget")
        layout = QVBoxLayout(self)
        layout.addWidget(SubtitleLabel("스케줄 리스트", self))
        self.table = TableWidget(self)
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(['시간', '작업 이름', 'URL', '관리'])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        
        # 더블클릭 시 셀 편집 방지
        self.table.setEditTriggers(TableWidget.EditTrigger.NoEditTriggers)
        
        layout.addWidget(self.table)
        self.update_table()

    def update_table(self):
        schedules = self.main_window.config['schedules']
        self.table.setRowCount(len(schedules))
        for i, sched in enumerate(schedules):
            self.table.setItem(i, 0, QTableWidgetItem(f"{sched['hour']:02d}:{sched['minute']:02d}"))
            self.table.setItem(i, 1, QTableWidgetItem(sched.get('task_name', 'Unnamed')))
            self.table.setItem(i, 2, QTableWidgetItem(sched.get('url', '')))
            btn_delete = PushButton("삭제 🗑️", self)
            btn_delete.clicked.connect(lambda checked, idx=i: self.delete_schedule(idx))
            self.table.setCellWidget(i, 3, btn_delete)

    def delete_schedule(self, index):
        msg = MessageBox("스케줄 삭제", "정말로 이 스케줄을 삭제하시겠습니까?", self.main_window)
        if msg.exec():
            if 0 <= index < len(self.main_window.config['schedules']):
                del self.main_window.config['schedules'][index]
                ConfigManager.save(self.main_window.config)
                self.update_table()

class SettingWidget(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        self.setObjectName("SettingWidget")
        layout = QVBoxLayout(self)
        layout.setSpacing(20)
        layout.addWidget(SubtitleLabel("기본 설정", self))
        startup_layout = QHBoxLayout()
        startup_layout.addWidget(BodyLabel("부팅 시 자동 실행:"))
        self.startup_switch = SwitchButton(self)
        self.startup_switch.setChecked(self.main_window.config.get('run_at_startup', False))
        self.startup_switch.checkedChanged.connect(self.toggle_startup)
        startup_layout.addWidget(self.startup_switch)
        startup_layout.addStretch(1)
        layout.addLayout(startup_layout)
        
        theme_layout = QHBoxLayout()
        theme_layout.addWidget(BodyLabel("다크 모드 사용:"))
        self.theme_switch = SwitchButton(self)
        self.theme_switch.setChecked(self.main_window.config.get('theme', 'Light') == 'Dark')
        self.theme_switch.checkedChanged.connect(self.toggle_theme)
        theme_layout.addWidget(self.theme_switch)
        theme_layout.addStretch(1)
        layout.addLayout(theme_layout)
        
        browser_layout = QHBoxLayout()
        browser_layout.addWidget(BodyLabel("기본 브라우저:"))
        self.browser_combo = ComboBox(self)
        self.browser_combo.addItems(["Chrome", "Edge"])
        self.browser_combo.setCurrentText(self.main_window.config.get('browser', 'Chrome'))
        self.browser_combo.currentTextChanged.connect(self.change_browser)
        browser_layout.addWidget(self.browser_combo)
        browser_layout.addStretch(1)
        layout.addLayout(browser_layout)

        # --- 바로가기 생성 섹션 추가 ---
        shortcut_layout = QHBoxLayout()
        shortcut_layout.addWidget(BodyLabel("바탕화면 바로가기:"))
        self.shortcut_button = PushButton("바로가기 만들기", self)
        self.shortcut_button.clicked.connect(self.create_shortcut)
        shortcut_layout.addWidget(self.shortcut_button)
        shortcut_layout.addStretch(1)
        layout.addLayout(shortcut_layout)

        # --- 구분선 및 제작자 정보 추가 ---
        layout.addSpacing(10)
        layout.addWidget(HorizontalSeparator(self))
        layout.addSpacing(10)
        
        creator_label = BodyLabel("오작동(버그)제보: phb@somunnanshop.com", self)
        creator_label.setStyleSheet("color: gray;")
        layout.addWidget(creator_label)
        
        layout.addStretch(1)
        
        # --- 최하단 중앙 버전 정보 ---
        version_label = CaptionLabel(f"Version {VERSION}", self)
        version_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(version_label)

    def toggle_startup(self, is_checked):
        success = RegistryUtils.set_startup(is_checked)
        if success:
            self.main_window.config['run_at_startup'] = is_checked
            ConfigManager.save(self.main_window.config)
        else:
            # 실패 시 스위치를 원상복구하고 에러 메시지 표시
            self.startup_switch.setChecked(not is_checked)
            InfoBar.error("설정 오류", "자동 실행 설정에 실패했습니다. 권한이 없거나 보안 프로그램에 의해 차단되었을 수 있습니다.", 
                          duration=3000, parent=self.main_window)

    def toggle_theme(self, is_checked):
        theme = "Dark" if is_checked else "Light"
        self.main_window.config['theme'] = theme
        ConfigManager.save(self.main_window.config)
        setTheme(Theme.DARK if is_checked else Theme.LIGHT)

    def change_browser(self, text):
        self.main_window.config['browser'] = text
        ConfigManager.save(self.main_window.config)

    def create_shortcut(self):
        """ 바탕화면 바로가기 생성 확인 메시지 및 실행 """
        msg = MessageBox("바로가기 생성", "바탕화면에 바로가기 아이콘을 생성 할까요?", self.main_window)
        msg.yesButton.setText("확인")
        msg.cancelButton.setText("취소")
        
        if msg.exec():
            if ShortcutUtils.create_desktop_shortcut():
                InfoBar.success("성공", "바탕화면에 바로가기가 생성되었습니다.", duration=2000, parent=self.main_window)
            else:
                InfoBar.error("실패", "바로가기 생성 중 오류가 발생했습니다.", duration=3000, parent=self.main_window)

class MainWindow(FluentWindow):
    def __init__(self):
        super().__init__()
        self.config = ConfigManager.load()
        setTheme(Theme.DARK if self.config.get('theme') == 'Dark' else Theme.LIGHT)
        self.setWindowTitle("ANU Auto Scheduler")
        self.resize(900, 700)

        # --- 아이콘 및 툴팁 설정 ---
        # 노란색(Gold) 아이콘 생성 (초시계와 유사한 HISTORY 아이콘 사용)
        from PyQt6.QtGui import QColor
        app_icon = FIF.HISTORY.icon(color=QColor("#FFD700")) # 금색/노란색
        self.setWindowIcon(app_icon)

        # 자동 실행 설정이 켜져 있으면 현재 위치로 레지스트리 정보 최신화
        if self.config.get('run_at_startup', False):
            RegistryUtils.set_startup(True)
        
        self.player = QMediaPlayer()
        self.audio_output = QAudioOutput()
        self.player.setAudioOutput(self.audio_output)
        self.audio_output.setVolume(0.8)
        
        self.home_widget = HomeWidget(self)
        self.setting_widget = SettingWidget(self)
        
        # 네비게이션 아이템 추가
        self.addSubInterface(self.home_widget, FIF.HOME, "스케줄 리스트")
        
        # 뒤로가기 버튼과 햄버거 메뉴 버튼 숨기기
        self.navigationInterface.setReturnButtonVisible(False)
        self.navigationInterface.setMenuButtonVisible(False)
        
        # '스케줄 추가'는 탭이 아닌 클릭 시 팝업을 띄우는 버튼으로 추가
        self.navigationInterface.addItem(
            routeKey='add_schedule',
            icon=FIF.ADD,
            text='스케줄 추가',
            onClick=self.show_add_schedule_dialog,
            position=NavigationItemPosition.TOP
        )
        
        self.addSubInterface(self.setting_widget, FIF.SETTING, "설정", NavigationItemPosition.BOTTOM)
        
        self.tray_icon = QSystemTrayIcon(self)
        self.tray_icon.setIcon(app_icon) # 트레이 아이콘도 동일하게 설정
        self.tray_icon.setToolTip(APP_NAME) # 마우스 오버 시 표시될 이름
        
        tray_menu = QMenu()
        show_action = QAction("열기", self)
        show_action.triggered.connect(self.show_window)
        quit_action = QAction("종료", self)
        quit_action.triggered.connect(QApplication.instance().quit)
        tray_menu.addAction(show_action)
        tray_menu.addAction(quit_action)
        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.show()
        self.tray_icon.activated.connect(self.tray_activated)
        
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.check_schedule)
        self.timer.start(1000)

    def show_add_schedule_dialog(self):
        """ 스케줄 추가 팝업 다이얼로그 표시 """
        dialog = AddScheduleDialog(self)
        dialog.exec()

    def show_window(self):
        if self.isMinimized():
            self.showNormal()
        self.show()
        self.raise_()
        self.activateWindow()
        
        # 윈도우에서 강제로 최상단으로 가져오기 (ctypes 사용)
        if sys.platform == "win32":
            hwnd = int(self.winId())
            ctypes.windll.user32.ShowWindow(hwnd, 9) # SW_RESTORE
            ctypes.windll.user32.SetForegroundWindow(hwnd)

    def tray_activated(self, reason):
        if reason == QSystemTrayIcon.ActivationReason.DoubleClick:
            self.show_window()

    def closeEvent(self, event):
        event.ignore()
        self.hide()
        InfoBar.info("백그라운드 실행", "프로그램이 시스템 트레이에서 계속 실행됩니다.", duration=2000, parent=self)

    def check_schedule(self):
        now = datetime.now()
        today_str = now.strftime("%Y-%m-%d")
        for sched in self.config['schedules']:
            if now.hour == sched['hour'] and now.minute == sched['minute']:
                if sched.get('last_run_date') != today_str:
                    sched['last_run_date'] = today_str
                    ConfigManager.save(self.config)
                    self.trigger_alarm(sched)

    def trigger_alarm(self, sched):
        if os.path.exists(SOUND_FILE):
            self.player.setSource(QUrl.fromLocalFile(SOUND_FILE))
            self.player.play()
        else:
            InfoBar.warning("알람 소리 누락", f"알람 파일({SOUND_FILE})을 찾을 수 없습니다.", duration=3000, parent=self)
        
        self.show_window()
        msg = MessageBox("스케줄 실행 알림", 
                         f"[{sched['task_name']}] 작업을 지금 실행하시겠습니까?\n\nURL: {sched['url']}", 
                         self)
        msg.yesButton.setText("실행")
        msg.cancelButton.setText("건너뛰기")
        if msg.exec():
            self.execute_task(sched)
        self.player.stop()

    def execute_task(self, sched):
        browser_pref = self.config.get('browser', 'Chrome')
        browser_path = RegistryUtils.get_browser_path(browser_pref)
        url = sched['url']
        if not url.startswith(('http://', 'https://')):
            url = 'https://' + url
        success = False
        if browser_path and os.path.exists(browser_path):
            try:
                subprocess.Popen([browser_path, url])
                success = True
            except Exception as e:
                print(f"Browser launch error: {e}")
        if not success:
            try:
                webbrowser.open(url)
            except Exception as e:
                InfoBar.error("실행 오류", f"URL을 열 수 없습니다: {e}", parent=self)

if __name__ == '__main__':
    # --- 작업 표시줄 아이콘 강제 설정 (AppUserModelID) ---
    if sys.platform == "win32":
        myappid = f'mycompany.myproduct.subproduct.{VERSION}' # 고유 ID
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

    QApplication.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    app = QApplication(sys.argv)
    
    # --- 중복 실행 방지 로직 ---
    socket = QLocalSocket()
    socket.connectToServer(SERVER_NAME)
    
    # 이미 서버가 존재하면 (프로그램이 실행 중이면)
    if socket.waitForConnected(500):
        # 기존 인스턴스에 메시지 전송 (연결만으로도 newConnection 이벤트 발생)
        socket.disconnectFromServer()
        sys.exit(0)
        
    # 서버가 없으면 (내가 첫 번째 인스턴스면) 서버 생성
    local_server = QLocalServer()
    if not local_server.listen(SERVER_NAME):
        # 만약 비정상 종료 등으로 소켓 파일이 남아있다면 삭제 후 재시도
        QLocalServer.removeServer(SERVER_NAME)
        local_server.listen(SERVER_NAME)

    window = MainWindow()
    
    # 새로운 연결 요청이 오면 (사용자가 프로그램을 또 실행하려고 하면) 창 표시
    local_server.newConnection.connect(window.show_window)
    
    window.show()
    sys.exit(app.exec())
