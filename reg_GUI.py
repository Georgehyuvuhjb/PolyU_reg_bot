import sys
import logging
import os
import time
import requests
import re
from urllib.parse import urljoin
from bs4 import BeautifulSoup

# Using PyQt6
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QLabel, QLineEdit, QPushButton, 
                             QTableWidget, QTableWidgetItem, QHeaderView, 
                             QTextEdit, QMessageBox, QGroupBox, QFormLayout)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QObject
from PyQt6.QtGui import QFont, QIcon

# ==========================================
# 1. Path Helper for PyInstaller
# ==========================================
def get_resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# ==========================================
# 2. Log Handler: Redirects Logger to GUI
# ==========================================
class SignallableLogHandler(logging.Handler, QObject):
    log_signal = pyqtSignal(str)

    def __init__(self):
        logging.Handler.__init__(self)
        QObject.__init__(self)

    def emit(self, record):
        msg = self.format(record)
        self.log_signal.emit(msg)

logger = logging.getLogger("PolyURegBot")
logger.setLevel(logging.INFO)
formatter = logging.Formatter('%(asctime)s - %(message)s')

# ==========================================
# 3. Core Logic: PolyU Course Reg System
# ==========================================
class CourseRegistrationSystem:
    def __init__(self, user_id, password):
        self.myid = user_id
        self.myPassword = password
        self.session = requests.Session()
        self.headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        }
        self.view_state = None
        
        self.base_url = "https://www38.polyu.edu.hk/eStudent/"
        self.acad_year_sem_url = urljoin(self.base_url, "secure/my-subject-registration/subject-register-select-acad-year-sem.jsf")
        self.subject_selection_url = urljoin(self.base_url, "secure/my-subject-registration/subject-register-select-subject.jsf")
        self.component_selection_url = urljoin(self.base_url, "secure/my-subject-registration/subject-register-select-component.jsf")
        self.preview_confirmation_url = urljoin(self.base_url, "secure/my-subject-registration/subject-register-preview-confirmation.jsf")

    def update_view_state(self, soup):
        vs = soup.find("input", {"name": "javax.faces.ViewState"})
        if vs:
            self.view_state = vs["value"]
            return True
        return False

    def login(self):
        try:
            logger.info("Step 1: Connecting to PolyU Auth Server...")
            start_url = "https://www38.polyu.edu.hk/eStudent/SAML_callback?eStudver=2"
            res = self.session.get(start_url, headers=self.headers)
            
            soup = BeautifulSoup(res.text, 'html.parser')
            form = soup.find('form', id='loginForm')
            if not form: return False
            
            post_url = urljoin(res.url, form.get('action'))
            payload = {tag.get('name'): tag.get('value', '') for tag in form.find_all('input') if tag.get('name')}
            payload['UserName'] = 'hh\\' + self.myid
            payload['Password'] = self.myPassword

            logger.info("Step 2: Submitting credentials for ADFS Authentication...")
            res = self.session.post(post_url, data=payload, headers=self.headers)
            
            if "SAMLResponse" not in res.text:
                logger.error("Login Failed: Check your ID and Password.")
                return False

            logger.info("Step 3: ADFS Authenticated. Returning to eStudent...")
            soup = BeautifulSoup(res.text, 'html.parser')
            saml_form = soup.find('form')
            sp_url = urljoin(res.url, saml_form.get('action'))
            saml_payload = {tag.get('name'): tag.get('value', '') for tag in saml_form.find_all('input') if tag.get('name')}
            
            final_res = self.session.post(sp_url, data=saml_payload, headers=self.headers)
            
            if final_res.ok:
                logger.info("Step 4: Login successful. Home page reached.")
                return True
            return False
        except Exception as e:
            logger.error(f"Login Error: {e}")
            return False

    def select_acad_year_sem(self):
        try:
            logger.info("Step 5: Accessing Academic Year/Semester selection...")
            res = self.session.get(self.acad_year_sem_url, headers=self.headers)
            self.update_view_state(BeautifulSoup(res.text, 'html.parser'))
            
            data = {"mainForm": "mainForm", "mainForm:nextButton": "Go", "javax.faces.ViewState": self.view_state}
            res = self.session.post(self.acad_year_sem_url, data=data, headers=self.headers)
            
            res = self.session.get(self.subject_selection_url, headers=self.headers)
            self.update_view_state(BeautifulSoup(res.text, 'html.parser'))
            return True
        except Exception as e:
            logger.error(f"Step 5 Failed: {e}")
            return False

    def add_subject(self, code, group, comps):
        try:
            logger.info(f"Processing: {code} (Group: {group})")
            data = {"mainForm": "mainForm", "mainForm:basicSearchSubjectCode": code, "mainForm:basicSearchButton": "Search", "javax.faces.ViewState": self.view_state}
            res = self.session.post(self.subject_selection_url, data=data, headers=self.headers)
            soup = BeautifulSoup(res.text, 'html.parser')
            self.update_view_state(soup)

            target_val = None
            select = soup.find("select", {"id": lambda x: x and "basicSearchSubjectGroup_" in x})
            if select:
                for opt in select.find_all("option"):
                    if group in opt.text:
                        target_val = opt.get('value')
                        break
            
            if not target_val:
                logger.warning(f"Group {group} not found for {code}")
                return False

            data = {"mainForm": "mainForm", "mainForm:basicSearchSubjectCode": code, f"mainForm:basicSearchTable:0:basicSearchSubjectGroup_": target_val, f"mainForm:basicSearchTable:0:basicSearchAddSubjectButton_": "+", "javax.faces.ViewState": self.view_state}
            res = self.session.post(self.subject_selection_url, data=data, headers=self.headers)
            
            comp_soup = BeautifulSoup(res.text, 'html.parser')
            self.update_view_state(comp_soup)
            
            comp_data = {"mainForm": "mainForm", "mainForm:selectCompSubjectGroup": target_val, "mainForm:selectButton": "Add to Cart", "javax.faces.ViewState": self.view_state}
            found = False
            for chk in comp_soup.find_all("input", {"type": "checkbox"}):
                row_text = chk.find_parent("tr").text
                for c in comps:
                    if c in row_text:
                        match = re.search(r':(\d+):selectCompSelected_', chk.get('id', ''))
                        if match:
                            comp_data[f"mainForm:ComponentTable:{match.group(1)}:selectCompSelected_"] = "on"
                            found = True
            
            if found:
                res = self.session.post(self.component_selection_url, data=comp_data, headers=self.headers)
                self.update_view_state(BeautifulSoup(res.text, 'html.parser'))
                logger.info(f"Successfully added {code} to cart.")
                return True
            return False
        except Exception as e:
            logger.error(f"Error adding {code}: {e}")
            return False

    def finalize(self):
        try:
            logger.info("Step 6: Confirming Shopping Cart...")
            data = {"mainForm": "mainForm", "mainForm:confirmButton": "Proceed to Preview", "javax.faces.ViewState": self.view_state}
            res = self.session.post(self.subject_selection_url, data=data, headers=self.headers)
            self.update_view_state(BeautifulSoup(res.text, 'html.parser'))
            
            data = {"mainForm": "mainForm", "mainForm:confirmButton": "Confirm", "javax.faces.ViewState": self.view_state}
            res = self.session.post(self.preview_confirmation_url, data=data, headers=self.headers)
            
            if "success" in res.text.lower() or "成功" in res.text:
                logger.info(">>> ALL TASKS COMPLETED SUCCESSFULLY! <<<")
                return True
            return False
        except Exception as e:
            logger.error(f"Final submission failed: {e}")
            return False

# ==========================================
# 4. Background Worker
# ==========================================
class Worker(QThread):
    finished_signal = pyqtSignal(bool, str)

    def __init__(self, uid, pwd, subjects):
        super().__init__()
        self.uid, self.pwd, self.subjects = uid, pwd, subjects

    def run(self):
        bot = CourseRegistrationSystem(self.uid, self.pwd)
        if not bot.login():
            self.finished_signal.emit(False, "Login failed.")
            return
        
        if not bot.select_acad_year_sem():
            self.finished_signal.emit(False, "Semester selection failed.")
            return

        success_any = False
        for s in self.subjects:
            if bot.add_subject(s[0], s[1], s[2]):
                success_any = True
            time.sleep(0.3)

        if success_any:
            if bot.finalize():
                self.finished_signal.emit(True, "Process finished successfully.")
            else:
                self.finished_signal.emit(False, "Failed to submit cart.")
        else:
            self.finished_signal.emit(False, "No subjects were added.")

# ==========================================
# 5. GUI Main Window
# ==========================================
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PolyU Registration Helper")
        self.resize(720, 600)
        
        # Set Window Icon (Local file)
        self.setWindowIcon(QIcon(get_resource_path("icon.ico")))
        
        self.setup_ui()
        
        self.log_handler = SignallableLogHandler()
        self.log_handler.setFormatter(formatter)
        self.log_handler.log_signal.connect(self.append_log)
        logger.addHandler(self.log_handler)

    def setup_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        # Login
        gb_login = QGroupBox("Login Credentials")
        fl = QFormLayout(gb_login)
        self.id_in = QLineEdit(); self.pw_in = QLineEdit()
        self.pw_in.setEchoMode(QLineEdit.EchoMode.Password)
        fl.addRow("Student ID:", self.id_in)
        fl.addRow("Password:", self.pw_in)
        layout.addWidget(gb_login)

        # Table
        gb_table = QGroupBox("Subject List")
        vl = QVBoxLayout(gb_table)
        self.table = QTableWidget(0, 3)
        self.table.setHorizontalHeaderLabels(["Code", "Group", "Components"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        
        hl = QHBoxLayout()
        btn_add = QPushButton("+ Add"); btn_add.clicked.connect(self.add_row)
        btn_del = QPushButton("- Delete"); btn_del.clicked.connect(self.del_row)
        hl.addWidget(btn_add); hl.addWidget(btn_del)
        
        vl.addWidget(self.table)
        vl.addLayout(hl)
        layout.addWidget(gb_table)

        # Start Button
        self.run_btn = QPushButton("RUN TASK")
        self.run_btn.setFixedHeight(50)
        self.run_btn.setStyleSheet("background-color: #900000; color: white; font-weight: bold; font-size: 14px; border-radius: 5px;")
        self.run_btn.clicked.connect(self.start)
        layout.addWidget(self.run_btn)

        # Logs
        self.log_out = QTextEdit(); self.log_out.setReadOnly(True)
        self.log_out.setStyleSheet("background-color: #1e1e1e; color: #ffffff; font-family: Consolas;")
        layout.addWidget(QLabel("Logs:"))
        layout.addWidget(self.log_out)

        self.add_example("ABCT1D04", "2001", "LTL001")

    def add_row(self): self.table.insertRow(self.table.rowCount())
    def add_example(self, c, g, m):
        r = self.table.rowCount(); self.table.insertRow(r)
        self.table.setItem(r, 0, QTableWidgetItem(c))
        self.table.setItem(r, 1, QTableWidgetItem(g))
        self.table.setItem(r, 2, QTableWidgetItem(m))

    def del_row(self): self.table.removeRow(self.table.currentRow())

    def append_log(self, text):
        self.log_out.append(text)
        self.log_out.verticalScrollBar().setValue(self.log_out.verticalScrollBar().maximum())

    def start(self):
        uid, pwd = self.id_in.text().strip(), self.pw_in.text().strip()
        if not uid or not pwd:
            QMessageBox.warning(self, "Warning", "Please enter ID and Password.")
            return

        subjects = []
        for i in range(self.table.rowCount()):
            c = self.table.item(i,0).text().strip() if self.table.item(i,0) else ""
            g = self.table.item(i,1).text().strip() if self.table.item(i,1) else ""
            m = self.table.item(i,2).text().strip() if self.table.item(i,2) else ""
            if c and g: subjects.append([c.upper(), g, [x.strip() for x in m.split(',')]])

        if not subjects:
            QMessageBox.warning(self, "Warning", "List is empty.")
            return

        self.run_btn.setEnabled(False); self.run_btn.setText("TASK RUNNING...")
        self.log_out.clear()
        logger.info("DISCLAIMER: This tool is for educational purposes only.")
        logger.info("-" * 40)
        
        self.worker = Worker(uid, pwd, subjects)
        self.worker.finished_signal.connect(self.on_done)
        self.worker.start()

    def on_done(self, ok, msg):
        self.run_btn.setEnabled(True); self.run_btn.setText("RUN TASK")
        if ok: QMessageBox.information(self, "Finished", msg)
        else: QMessageBox.critical(self, "Error", msg)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setFont(QFont("Segoe UI", 11))
    win = MainWindow()
    win.show()
    sys.exit(app.exec())