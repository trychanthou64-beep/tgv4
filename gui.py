import sys
import os
import warnings
warnings.filterwarnings("ignore", category=DeprecationWarning)
os.environ.pop("_MEIPASS2", None)
os.environ["QT_LOGGING_RULES"] = "*.warning=false"
import json
import subprocess
from PyQt6.QtCore import QProcess, Qt, QFileSystemWatcher, QProcessEnvironment
from PyQt6.QtGui import QFont, QIcon, QColor, QPalette
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QPlainTextEdit, QFrame,
    QMessageBox, QGraphicsDropShadowEffect
)

if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))


class ModernBotControlPanel(QMainWindow):
    def __init__(self):
        super().__init__()
        self.config_file = os.path.join(BASE_DIR, 'config.json')
        self.default_token = "8788148855:AAG3f3N_vCMZ7BEo1hjCcK721zIWRyij4pI"
        self.default_admin = [8558847170]
        self.default_menus = ["F88$", "F88$_CH", "F88_CH_KH", "F88_KH"]
        self.process = None

        self.init_ui()
        self.load_config()

        # Monitor config.json for dynamic updates
        self.watcher = QFileSystemWatcher()
        if os.path.exists(self.config_file):
            self.watcher.addPath(self.config_file)
        self.watcher.fileChanged.connect(self.on_config_file_changed)

    def init_ui(self):
        self.setWindowTitle("Telegram Bot Controller Pro (ប្រព័ន្ធគ្រប់គ្រង Telegram Bot)")
        self.resize(820, 640)
        self.setMinimumSize(720, 540)

        # Set Window Icon
        if getattr(sys, 'frozen', False):
            icon_dir = getattr(sys, '_MEIPASS', BASE_DIR)
        else:
            icon_dir = BASE_DIR
        icon_path = os.path.join(icon_dir, 'app_icon.png')
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        self.apply_theme()

        # Central Widget & Main Layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(24, 20, 24, 24)
        main_layout.setSpacing(18)

        # ---------------- 1. HEADER BAR ----------------
        header_card = QFrame()
        header_card.setObjectName("HeaderCard")
        header_layout = QHBoxLayout(header_card)
        header_layout.setContentsMargins(18, 14, 18, 14)
        header_layout.setSpacing(15)

        # Header Title Stack
        title_box = QVBoxLayout()
        title_box.setSpacing(3)

        app_title = QLabel("⚡ TELEGRAM BOT CONTROLLER")
        app_title.setFont(QFont("Segoe UI", 15, QFont.Weight.Bold))
        app_title.setStyleSheet("color: #0F172A; letter-spacing: 1px;")

        app_subtitle = QLabel("Automated Dispatcher & Management Server | ប្រព័ន្ធគ្រប់គ្រង Telegram Bot ស្វ័យប្រវត្តិ")
        app_subtitle.setFont(QFont("Segoe UI", 9))
        app_subtitle.setStyleSheet("color: #64748B;")

        title_box.addWidget(app_title)
        title_box.addWidget(app_subtitle)
        header_layout.addLayout(title_box)

        header_layout.addStretch()

        # Live Status Badge Pill
        self.status_pill = QLabel("● STOPPED (បានបិទ)")
        self.status_pill.setObjectName("StatusPillStopped")
        self.status_pill.setFont(QFont("Segoe UI", 9, QFont.Weight.Bold))
        self.status_pill.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.status_pill.setFixedSize(165, 34)
        header_layout.addWidget(self.status_pill)

        main_layout.addWidget(header_card)

        # ---------------- 2. CONFIGURATION CARD ----------------
        config_card = QFrame()
        config_card.setObjectName("ConfigCard")
        config_layout = QVBoxLayout(config_card)
        config_layout.setContentsMargins(20, 18, 20, 18)
        config_layout.setSpacing(14)

        card_title = QLabel("🔑 BOT AUTHENTICATION & CONFIG | ការកំណត់ BOT")
        card_title.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
        card_title.setStyleSheet("color: #2563EB;")
        config_layout.addWidget(card_title)

        # Token Input Row
        token_row = QHBoxLayout()
        token_row.setSpacing(10)

        token_label = QLabel("Bot Token (កូដ Token):")
        token_label.setFont(QFont("Segoe UI", 10, QFont.Weight.Medium))
        token_label.setStyleSheet("color: #334155;")
        token_row.addWidget(token_label)

        self.token_input = QLineEdit()
        self.token_input.setPlaceholderText("Paste Telegram Bot Token (ឧ. 123456789:ABCDefgh...)")
        self.token_input.setEchoMode(QLineEdit.EchoMode.Password)
        self.token_input.setFont(QFont("Consolas", 10))
        self.token_input.setObjectName("TokenInput")
        token_row.addWidget(self.token_input)

        self.toggle_token_btn = QPushButton("👁️ Show (បង្ហាញ)")
        self.toggle_token_btn.setCheckable(True)
        self.toggle_token_btn.setFont(QFont("Segoe UI", 9, QFont.Weight.Bold))
        self.toggle_token_btn.setObjectName("SecondaryBtn")
        self.toggle_token_btn.setFixedWidth(120)
        self.toggle_token_btn.clicked.connect(self.toggle_token_visibility)
        token_row.addWidget(self.toggle_token_btn)

        config_layout.addLayout(token_row)

        # Save Configuration Button
        self.save_btn = QPushButton("💾 Save Configuration (រក្សាទុកការកំណត់)")
        self.save_btn.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
        self.save_btn.setObjectName("PrimaryBtn")
        self.save_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.save_btn.clicked.connect(self.save_config)
        config_layout.addWidget(self.save_btn)

        main_layout.addWidget(config_card)

        # ---------------- 3. CONTROLS TOOLBAR ----------------
        control_card = QFrame()
        control_card.setObjectName("ControlCard")
        control_layout = QHBoxLayout(control_card)
        control_layout.setContentsMargins(18, 12, 18, 12)
        control_layout.setSpacing(14)

        self.start_btn = QPushButton("▶ START BOT (ចាប់ផ្តើម)")
        self.start_btn.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
        self.start_btn.setObjectName("StartBtn")
        self.start_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.start_btn.clicked.connect(self.start_bot)
        control_layout.addWidget(self.start_btn)

        self.stop_btn = QPushButton("⏹ STOP BOT (បញ្ឈប់)")
        self.stop_btn.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
        self.stop_btn.setObjectName("StopBtn")
        self.stop_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.stop_btn.setEnabled(False)
        self.stop_btn.clicked.connect(self.stop_bot)
        control_layout.addWidget(self.stop_btn)

        control_layout.addStretch()

        self.clear_log_btn = QPushButton("🧹 Clear Log (លុប Log)")
        self.clear_log_btn.setFont(QFont("Segoe UI", 9, QFont.Weight.Medium))
        self.clear_log_btn.setObjectName("SecondaryBtn")
        self.clear_log_btn.clicked.connect(self.clear_console_log)
        control_layout.addWidget(self.clear_log_btn)

        self.copy_log_btn = QPushButton("📋 Copy Log (ចម្លង Log)")
        self.copy_log_btn.setFont(QFont("Segoe UI", 9, QFont.Weight.Medium))
        self.copy_log_btn.setObjectName("SecondaryBtn")
        self.copy_log_btn.clicked.connect(self.copy_console_log)
        control_layout.addWidget(self.copy_log_btn)

        main_layout.addWidget(control_card)

        # ---------------- 4. CONSOLE OUTPUT ----------------
        console_card = QFrame()
        console_card.setObjectName("ConsoleCard")
        console_layout = QVBoxLayout(console_card)
        console_layout.setContentsMargins(18, 14, 18, 14)
        console_layout.setSpacing(10)

        console_title_row = QHBoxLayout()
        console_title = QLabel("💻 REAL-TIME CONSOLE OUTPUT | កំណត់ហេតុដំណើការ")
        console_title.setFont(QFont("Segoe UI", 9, QFont.Weight.Bold))
        console_title.setStyleSheet("color: #64748B; letter-spacing: 0.5px;")
        console_title_row.addWidget(console_title)
        console_title_row.addStretch()

        console_layout.addLayout(console_title_row)

        self.log_output = QPlainTextEdit()
        self.log_output.setReadOnly(True)
        self.log_output.setFont(QFont("Consolas", 10))
        self.log_output.setObjectName("ConsoleOutput")
        console_layout.addWidget(self.log_output)

        main_layout.addWidget(console_card)

    def apply_theme(self):
        self.setStyleSheet("""
            QMainWindow {
                background-color: #F1F5F9;
            }
            QWidget {
                color: #0F172A;
            }

            /* Card Containers */
            #HeaderCard, #ConfigCard, #ControlCard, #ConsoleCard {
                background-color: #FFFFFF;
                border: 1px solid #E2E8F0;
                border-radius: 12px;
            }

            /* Inputs */
            #TokenInput {
                background-color: #F8FAFC;
                color: #0F172A;
                border: 1px solid #CBD5E1;
                border-radius: 8px;
                padding: 9px 14px;
            }
            #TokenInput:focus {
                border: 1px solid #2563EB;
                background-color: #FFFFFF;
            }
            #TokenInput:disabled {
                background-color: #F1F5F9;
                color: #94A3B8;
                border: 1px solid #E2E8F0;
            }

            /* Buttons */
            #PrimaryBtn {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #2563EB, stop:1 #4F46E5);
                color: #FFFFFF;
                border: none;
                border-radius: 8px;
                padding: 10px 18px;
            }
            #PrimaryBtn:hover {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #1D4ED8, stop:1 #4338CA);
            }
            #PrimaryBtn:pressed {
                background-color: #1E40AF;
            }
            #PrimaryBtn:disabled {
                background-color: #E2E8F0;
                color: #94A3B8;
                border: none;
            }

            #SecondaryBtn {
                background-color: #F1F5F9;
                color: #334155;
                border: 1px solid #CBD5E1;
                border-radius: 8px;
                padding: 8px 14px;
            }
            #SecondaryBtn:hover {
                background-color: #E2E8F0;
                color: #0F172A;
            }
            #SecondaryBtn:pressed {
                background-color: #CBD5E1;
            }

            #StartBtn {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #10B981, stop:1 #059669);
                color: #FFFFFF;
                border: none;
                border-radius: 8px;
                padding: 12px 24px;
            }
            #StartBtn:hover {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #059669, stop:1 #047857);
            }
            #StartBtn:disabled {
                background-color: #E2E8F0;
                color: #94A3B8;
            }

            #StopBtn {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #EF4444, stop:1 #DC2626);
                color: #FFFFFF;
                border: none;
                border-radius: 8px;
                padding: 12px 24px;
            }
            #StopBtn:hover {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #DC2626, stop:1 #B91C1C);
            }
            #StopBtn:disabled {
                background-color: #E2E8F0;
                color: #94A3B8;
            }

            /* Status Pills */
            #StatusPillStopped {
                background-color: #FEE2E2;
                color: #DC2626;
                border: 1px solid #FCA5A5;
                border-radius: 17px;
            }
            #StatusPillRunning {
                background-color: #D1FAE5;
                color: #059669;
                border: 1px solid #6EE7B7;
                border-radius: 17px;
            }

            /* Console Text Area */
            #ConsoleOutput {
                background-color: #F8FAFC;
                color: #0F172A;
                border: 1px solid #CBD5E1;
                border-radius: 8px;
                padding: 12px;
            }
        """)

    def toggle_token_visibility(self, checked):
        if checked:
            self.token_input.setEchoMode(QLineEdit.EchoMode.Normal)
            self.toggle_token_btn.setText("🔒 Hide (លាក់)")
        else:
            self.token_input.setEchoMode(QLineEdit.EchoMode.Password)
            self.toggle_token_btn.setText("👁️ Show (បង្ហាញ)")

    def load_config(self):
        self.main_menus_data = {"Default": list(self.default_menus)}
        self.admin_ids = list(self.default_admin)
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                self.token_input.setText(config.get("TOKEN", self.default_token))
                menus = config.get("MENUS", self.default_menus)
                main_menus = config.get("MAIN_MENUS", {})
                if main_menus:
                    self.main_menus_data = main_menus
                elif menus:
                    self.main_menus_data = {"Default": list(menus)}
                self.admin_ids = config.get("ADMIN_IDS", self.default_admin)
            except Exception as e:
                self.log_output.appendPlainText(f"⚠️ Error loading configuration: {e}")
        else:
            self.token_input.setText(self.default_token)
            self.main_menus_data = {"Default": list(self.default_menus)}
            self.admin_ids = list(self.default_admin)

    def save_config(self):
        token = self.token_input.text().strip()
        menus = []
        for subs in self.main_menus_data.values():
            menus.extend(subs)
        if not menus:
            menus = list(self.default_menus)

        config_data = {
            "TOKEN": token,
            "MENUS": menus,
            "MAIN_MENUS": self.main_menus_data,
            "ADMIN_IDS": self.admin_ids
        }

        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config_data, f, indent=2)
            self.log_output.appendPlainText("✅ Configuration saved successfully to config.json. (បានរក្សាទុកការកំណត់ដោយជោគជ័យ)")
        except Exception as e:
            self.log_output.appendPlainText(f"❌ Error saving configuration: {e}")

    def on_config_file_changed(self, path):
        # Dynamically reload config when external changes happen
        self.load_config()

    def start_bot(self):
        if self.process is not None:
            return

        self.log_output.appendPlainText("🚀 Starting bot process... (កំពុងចាប់ផ្តើមប្រព័ន្ធ Bot...)")

        # Clean up any orphan bot.exe instances on Windows
        if sys.platform == "win32":
            subprocess.run("taskkill /f /im bot.exe", shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

        self.process = QProcess()
        self.process.setProcessChannelMode(QProcess.ProcessChannelMode.MergedChannels)

        env = QProcessEnvironment.systemEnvironment()
        env.insert("PYTHONUNBUFFERED", "1")
        env.insert("PYTHONIOENCODING", "UTF-8")
        env.insert("USER_BASE_DIR", BASE_DIR)
        if env.contains("_MEIPASS2"):
            env.remove("_MEIPASS2")
        if env.contains("_MEIPASS"):
            env.remove("_MEIPASS")
        self.process.setProcessEnvironment(env)

        self.process.readyRead.connect(self.read_bot_output)
        self.process.finished.connect(self.bot_finished)

        if getattr(sys, 'frozen', False):
            bot_executable = os.path.join(getattr(sys, '_MEIPASS', BASE_DIR), 'bot.exe')
            if not os.path.exists(bot_executable):
                bot_executable = os.path.join(os.path.dirname(sys.executable), 'bot.exe')
            if not os.path.exists(bot_executable):
                bot_executable = os.path.join(BASE_DIR, 'bot.exe')
            self.process.start(bot_executable)
        else:
            bot_script = os.path.join(BASE_DIR, 'bot.py')
            self.process.start(sys.executable, [bot_script])

        if self.process.waitForStarted(2500):
            self.status_pill.setText("● RUNNING (កំពុងដំណើការ)")
            self.status_pill.setObjectName("StatusPillRunning")
            self.status_pill.setStyle(self.status_pill.style())
            self.start_btn.setEnabled(False)
            self.stop_btn.setEnabled(True)
            self.token_input.setEnabled(False)
            self.save_btn.setEnabled(False)
        else:
            self.log_output.appendPlainText("❌ Failed to start bot script process.")
            self.process = None

    def read_bot_output(self):
        if self.process:
            data = self.process.readAll().data()
            try:
                text = data.decode('utf-8')
            except UnicodeDecodeError:
                text = data.decode('cp1252', errors='replace')
            self.log_output.insertPlainText(text)
            self.log_output.ensureCursorVisible()

    def stop_bot(self):
        if self.process:
            self.log_output.appendPlainText("🛑 Stopping bot process... (កំពុងបញ្ឈប់ប្រព័ន្ធ Bot...)")
            pid = self.process.processId()
            if pid and sys.platform == "win32":
                subprocess.run(f"taskkill /f /t /pid {pid}", shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                subprocess.run("taskkill /f /im bot.exe", shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            else:
                self.process.terminate()
                if not self.process.waitForFinished(3000):
                    self.process.kill()

    def bot_finished(self, exit_code, exit_status):
        self.log_output.appendPlainText(f"\n⏹ Bot process finished with exit code {exit_code}.")
        self.process = None
        self.status_pill.setText("● STOPPED (បានបិទ)")
        self.status_pill.setObjectName("StatusPillStopped")
        self.status_pill.setStyle(self.status_pill.style())
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.token_input.setEnabled(True)
        self.save_btn.setEnabled(True)

    def clear_console_log(self):
        self.log_output.clear()

    def copy_console_log(self):
        QApplication.clipboard().setText(self.log_output.toPlainText())
        self.log_output.appendPlainText("📋 Console log copied to clipboard.")

    def closeEvent(self, event):
        if self.process:
            self.stop_bot()
        event.accept()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ModernBotControlPanel()
    window.show()
    sys.exit(app.exec())
