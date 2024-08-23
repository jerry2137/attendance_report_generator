import sys
import os
import json
import datetime
import re
from base64 import b64encode, b64decode
from email.mime.text import MIMEText
import smtplib

from PyQt6 import QtGui
from PyQt6.QtCore import Qt, QTimer
from PyQt6.QtWidgets import (
    QApplication,
    QCheckBox,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
    QHBoxLayout,
    QScrollArea,
    QVBoxLayout,
    QWidget,
    QAbstractButton,
)

CWD = os.path.dirname(os.path.realpath(sys.argv[0]))
CONFIG_PATH = f'{CWD}/config.json'
PASSWORD_ICON_PATH = f'{CWD}/view_password.ico'

CHINESE_REASONS = ['休假', '上午休假', '下午休假', '病假', '出差', '在家工作', '夜班',]
ENGLISH_REASONS = ['leave', 'morning_leave', 'afternoon_leave', 'sick', 'business', 'home', 'night',]

SMTP_SERVER = 'email.cathaysec.com.tw'
REGEX_EMAIL = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,7}\b'

class SelectArea(QScrollArea):
    def __init__(self, master):
        super().__init__(widgetResizable=True)
        self.master = master
        temp_widget = QWidget()
        self.setWidget(temp_widget)
        self.vbox = QVBoxLayout(temp_widget)
        self.vbox.setAlignment(Qt.AlignmentFlag.AlignTop)

        self.items = {}

    def delete_item(self):
        item_name = self.sender().objectName()
        comfirm = QMessageBox.question(self, 'Warning', f'確認刪除：{item_name}？')
        if comfirm == QMessageBox.StandardButton.No:
            return
        hbox = self.findChild(QHBoxLayout, item_name)

        while hbox.count():
            box = hbox.takeAt(0).widget()
            box.setParent(None)
        hbox.setParent(None)

        del self.items[item_name]

    def add_item(self, item_name=None, data=None):
        pass
        
    def get_items(self):
        return self.items
        
    def set_items(self, items):
        for item_name, data in items.items():
            self.add_item(item_name, data)

class RecipientArea(SelectArea):
    def __init__(self, master, tip=''):
        super().__init__(master)
        
        self.add_hbox = QHBoxLayout()
        self.vbox.addLayout(self.add_hbox)

        self.input_box = QLineEdit(placeholderText=tip)
        self.input_box.setToolTip(tip)
        self.add_hbox.addWidget(self.input_box)

        add_button = QPushButton('add')
        add_button.setFixedWidth(100)
        add_button.clicked.connect(self.add_item)
        self.add_hbox.addWidget(add_button)

    def add_item(self, item_name=None, data=None):
        if not item_name:
            item_name = self.input_box.text()

        if not re.fullmatch(REGEX_EMAIL, item_name):
            QMessageBox.critical(self, 'Error', '請輸入email')
            return

        if item_name in self.items:
            QMessageBox.critical(self, 'Error', 'email重複')
            return
        
        self.items[item_name] = None
        
        hbox = QHBoxLayout(objectName=item_name)
        self.vbox.addLayout(hbox)

        hbox.addWidget(QLabel(item_name))
        
        delete_button = QPushButton('delete', objectName=item_name)
        delete_button.setFixedWidth(100)
        delete_button.clicked.connect(self.delete_item)
        hbox.addWidget(delete_button)

        self.input_box.setText('')

class NamesArea(SelectArea):
    def __init__(self, master):
        super().__init__(master)
        
        self.add_hbox = QHBoxLayout()
        self.vbox.addLayout(self.add_hbox)

        chinese_tip = '中文全名'
        self.chinese_box = QLineEdit(placeholderText=chinese_tip)
        self.chinese_box.setToolTip(chinese_tip)
        self.add_hbox.addWidget(self.chinese_box)
        
        english_tip = 'preferred English name'
        self.english_box = QLineEdit(placeholderText=english_tip)
        self.english_box.setToolTip(english_tip)
        self.add_hbox.addWidget(self.english_box)

        add_button = QPushButton('add')
        add_button.clicked.connect(self.add_item)
        self.add_hbox.addWidget(add_button)

    def add_item(self, chinese_name=None, english_name=None):
        if (not chinese_name) | (not english_name):
            chinese_name = self.chinese_box.text()
            english_name = self.english_box.text()

        if (not chinese_name) | (not english_name):
            QMessageBox.critical(self, 'Error', '必須同時輸入中文及英文名字')
            return

        if chinese_name in self.items:
            QMessageBox.critical(self, 'Error', f'不能輸入重複中文名: {chinese_name}')
            return
        
        hbox = QHBoxLayout(objectName=chinese_name)
        self.vbox.addLayout(hbox)

        # name
        hbox.addWidget(QLabel(chinese_name))
        hbox.addWidget(QLabel(english_name))

        # reasons
        for chinese_reason in CHINESE_REASONS:
            hbox.addWidget(QCheckBox(chinese_reason))
        
        delete_button = QPushButton('delete', objectName=chinese_name)
        delete_button.clicked.connect(self.delete_item)
        hbox.addWidget(delete_button)

        self.chinese_box.setText('')
        self.english_box.setText('')

        self.items[chinese_name] = english_name
        
    def get_attendance(self):
        attendance = {}
        attendance['all'] = list(self.items.values())
        for english_reason in ENGLISH_REASONS:
            attendance[english_reason] = []
        attendance['work'] = []

        for chinese_name, english_name in self.items.items():
            hbox = self.findChild(QHBoxLayout, chinese_name)
            will_work = True
            for i, english_reason in enumerate(ENGLISH_REASONS):
                if hbox.itemAt(i+2).widget().isChecked():
                    attendance[english_reason].append(english_name)
                    if english_reason not in ['morning_leave', 'afternoon_leave']:
                        will_work = False
            if will_work:
                attendance['work'].append(english_name)
        return attendance

class Window(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('出席報告產生器')
        
        self.center_widget = QWidget()
        self.setCentralWidget(self.center_widget)
        self.hbox = QHBoxLayout(self.center_widget)

        name_vbox = QVBoxLayout()
        self.hbox.addLayout(name_vbox)

        # Title
        self.title = QLabel('出席報告產生器', font=QtGui.QFont('Arial', 30))
        name_vbox.addWidget(self.title, alignment=Qt.AlignmentFlag.AlignCenter)

        # names
        self.read_folder_frame = NamesArea(self)
        name_vbox.addWidget(self.read_folder_frame)
        
        generate_tip = '產生郵件'
        self.generate_button = QPushButton(generate_tip)
        self.generate_button.setToolTip(generate_tip)
        self.generate_button.clicked.connect(self.generate)
        name_vbox.addWidget(self.generate_button)

        email_vbox = QVBoxLayout()
        self.hbox.addLayout(email_vbox)

        sender_tip = '寄件人'
        self.sender_box = QLineEdit(placeholderText=sender_tip)
        self.sender_box.setToolTip(sender_tip)
        email_vbox.addWidget(self.sender_box)

        password_tip = '密碼'
        self.password_box = QLineEdit(placeholderText=password_tip)
        self.password_box.setToolTip(password_tip)
        email_vbox.addWidget(self.password_box)
        self.set_password_box(self.password_box)
        
        recipient_tip = '收件人'
        self.recipient_frame = RecipientArea(self, tip=recipient_tip)
        email_vbox.addWidget(self.recipient_frame)

        header_tip = '郵件開頭'
        self.header_box = QTextEdit(placeholderText=header_tip)
        self.header_box.setToolTip(header_tip)
        email_vbox.addWidget(self.header_box)

        content_tip = '郵件內容'
        self.content_box = QTextEdit(placeholderText=content_tip)
        self.content_box.setToolTip(content_tip)
        email_vbox.addWidget(self.content_box)

        footer_tip = '郵件結尾'
        self.footer_box = QTextEdit(placeholderText=footer_tip)
        self.footer_box.setToolTip(footer_tip)
        email_vbox.addWidget(self.footer_box)
        
        send_tip = '傳送郵件'
        self.send_button = QPushButton(send_tip)
        self.send_button.setToolTip(send_tip)
        self.send_button.clicked.connect(self.send_email)
        email_vbox.addWidget(self.send_button)
        
        # Config
        self.read_config()

        # Start window
        self.showMaximized() 

    def read_config(self):
        try:
            with open(CONFIG_PATH, 'r') as file:
                self.config = json.load(file)
        except:
            QMessageBox.warning(self, 'Warning', '無法載入名單，請重新輸入')
            self.config = {}
            return

        self.read_folder_frame.set_items(self.config['names'])
        self.sender_box.setText(self.config['sender'])
        self.password_box.setText(b64decode(self.config['password_encoded'].encode()).decode())
        self.recipient_frame.set_items(self.config['recipients'])
        self.header_box.setText(self.config['header'])
        self.footer_box.setText(self.config['footer'])

    def generate(self):
        attendence_dict = self.read_folder_frame.get_attendance()
        all_count = len(attendence_dict['all'])
        work_count = len(attendence_dict['work'])

        attendence_rows = [datetime.date.today().strftime('%Y年%#m月%#d日\r\n')]
        summary_rows = [f'同仁共{all_count}名：',]
        
        for chinese_reason, english_reason in zip(CHINESE_REASONS, ENGLISH_REASONS):
            if not attendence_dict[english_reason]:
                continue
            reason_count = len(attendence_dict[english_reason])
            attendence_rows.append(f'{chinese_reason}：{', '.join(attendence_dict[english_reason])}\r\n')
            summary_rows.append(f'{reason_count}名{chinese_reason}，')

        summary_rows.append(f'上班同仁{work_count}名',)
        attendence_report = ''.join(attendence_rows + summary_rows)

        self.content_box.setText(attendence_report)
        clipboard = QApplication.clipboard()
        clipboard.setText(attendence_report)
        self.show_message('Successful', '已複製到剪貼簿')

    def send_email(self):
        header = self.header_box.toPlainText()
        content = self.content_box.toPlainText()
        footer = self.footer_box.toPlainText()
        if not content:
            QMessageBox.critical(self, 'Error', '無內容，請先產生郵件')
            return
        email_text = '\r\n\r\n'.join([header, content, footer])
        
        sender = self.sender_box.text()
        password = self.password_box.text()
        recipients = self.recipient_frame.get_items()

        if not re.fullmatch(REGEX_EMAIL, sender):
            QMessageBox.critical(self, 'Error', '寄件人Email錯誤')
            return
        if not password:
            QMessageBox.critical(self, 'Error', '請輸入密碼')
            return
        if not recipients:
            QMessageBox.critical(self, 'Error', '無收件人')
            return

        message = MIMEText(email_text, 'plain', 'utf-8')
        message['From'] = 'Attendance Report <@>'
        message['To'] = ', '.join(recipients)
        message['Subject'] = datetime.date.today().strftime('Attendance Report %Y年%#m月%#d日')
        
        try:
            smtp = smtplib.SMTP(SMTP_SERVER, port=25, timeout=10)
        except Exception as e:
            QMessageBox.critical(self, 'Error', f'SMTP連線失敗，錯誤訊息：{e}')
            return

        with smtp:
            try:
                smtp.login(sender, password)
            except smtplib.SMTPAuthenticationError:
                QMessageBox.critical(self, 'Error', f'Email密碼錯誤')
                return
            except Exception as e:
                QMessageBox.critical(self, 'Error', f'Email登入失敗，錯誤訊息：{e}')
                return
            
            try:
                status = smtp.send_message(message)
            except Exception as e:
                QMessageBox.critical(self, 'Error', f'Email送出失敗，錯誤訊息：{e}')
                return
            if status:
                QMessageBox.critical(self, 'Error', f'Email送出失敗，錯誤狀態：{status}')
                return
            self.show_message('Successful', 'Email已送出')

    def set_password_box(self, password_box):
        showPassAction = QtGui.QAction(QtGui.QIcon(PASSWORD_ICON_PATH), 'Show password', self)
        password_box.addAction(showPassAction, QLineEdit.ActionPosition.TrailingPosition)
        button = password_box.findChild(QAbstractButton)
        button.pressed.connect(lambda: password_box.setEchoMode(QLineEdit.EchoMode.Normal))
        button.released.connect(lambda: password_box.setEchoMode(QLineEdit.EchoMode.Password))
        password_box.setEchoMode(QLineEdit.EchoMode.Password)

    def show_message(self, title, message):
        messagebox = QMessageBox(windowTitle=title, text=message, icon=QMessageBox.Icon.Information)
        messagebox.setStandardButtons(QMessageBox.StandardButton.NoButton)
        messagebox.setWindowFlags(Qt.WindowType.FramelessWindowHint)
        QTimer.singleShot(500, messagebox.accept)
        messagebox.exec()

    def closeEvent(self, event):
        self.config['names'] = self.read_folder_frame.get_items()
        self.config['sender'] = self.sender_box.text()
        self.config['password_encoded'] = b64encode(self.password_box.text().encode()).decode()
        self.config['recipients'] = self.recipient_frame.get_items()
        self.config['header'] = self.header_box.toPlainText()
        self.config['footer'] = self.footer_box.toPlainText()
        with open(CONFIG_PATH, 'w') as file:
            json.dump(self.config, file, indent=4)
        QApplication.closeAllWindows()
        event.accept()

def start():
    app = QApplication(sys.argv)
    window = Window()
    sys.exit(app.exec())


if __name__ == '__main__':
    start()