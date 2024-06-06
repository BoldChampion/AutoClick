import sys, requests, ddddocr
import os
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLineEdit,
                             QTreeView, QMessageBox, QMenu, QFileDialog, QInputDialog, QComboBox, QStyledItemDelegate, QTextEdit)
from PyQt5.QtGui import QStandardItemModel, QStandardItem
from PyQt5.QtCore import Qt, QThread, pyqtSignal, pyqtSlot
import pandas as pd
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException, WebDriverException, NoSuchWindowException
import time
import pyautogui

class SeleniumWorker(QThread):
    finished = pyqtSignal()
    driver_created = pyqtSignal(object)  # Signal to send the driver instance
    error_signal = pyqtSignal(str)

    def __init__(self, url, model, driver=None):
        super().__init__()
        self.driver = driver
        self.url = url
        self.model = model
        self.running = True

    def run(self):
        try:
            if self.driver is None or not self.driver.session_id:
                options = Options()
                options.use_chromium = True
                service = Service(executable_path="msedgedriver.exe")  # Replace with actual path to msedgedriver
                self.driver = webdriver.Edge(service=service, options=options)
                self.driver_created.emit(self.driver)  # Emit the signal with the driver instance
            try:
                self.driver.get(self.url)
            except WebDriverException as e:
                self.error_signal.emit(f"WebDriver error: {str(e)}")
                return

            wait = WebDriverWait(self.driver, 30)  # Set explicit wait time to 10 seconds

            for row in range(self.model.rowCount()):
                if not self.running:
                    break
                mode = self.model.item(row, 0).text()
                path = self.model.item(row, 1).text()
                data = self.model.item(row, 2).text()
                if path != "Enter Path":
                    try:
                        match mode:
                            case "url":
                                self.driver.get(path)
                            case "iframein":
                                frame_element = wait.until(EC.visibility_of_element_located((By.XPATH, path)))
                                self.driver.switch_to.frame(frame_element)
                            case "iframeout":
                                self.driver.switch_to.default_content()
                            case "id":
                                element = wait.until(EC.visibility_of_element_located((By.ID, path)))
                                self.handle_element(element, data)
                            case "name":
                                element = wait.until(EC.visibility_of_element_located((By.NAME, path)))
                                self.handle_element(element, data)
                            case "class_name":
                                element = wait.until(EC.visibility_of_element_located((By.CLASS_NAME, path)))
                                self.handle_element(element, data)
                            case "tag":
                                element = wait.until(EC.visibility_of_element_located((By.TAG_NAME, path)))
                                self.handle_element(element, data)
                            case "link_text":
                                element = wait.until(EC.visibility_of_element_located((By.LINK_TEXT, path)))
                                self.handle_element(element, data)
                            case "partial_link_text":
                                element = wait.until(EC.visibility_of_element_located((By.PARTIAL_LINK_TEXT, path)))
                                self.handle_element(element, data)
                            case "css_selector":
                                element = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, path)))
                                self.handle_element(element, data)
                            case "xpath":
                                element = wait.until(EC.visibility_of_element_located((By.XPATH, path)))
                                self.handle_element(element, data)
                            case "c_id":
                                element = wait.until(EC.visibility_of_element_located((By.ID, path)))
                                data = self.getCaptcha(data)
                                self.handle_element(element, data)
                            case "c_name":
                                element = wait.until(EC.visibility_of_element_located((By.NAME, path)))
                                data = self.getCaptcha(data)
                                self.handle_element(element, data)
                            case "c_class_name":
                                element = wait.until(EC.visibility_of_element_located((By.CLASS_NAME, path)))
                                data = self.getCaptcha(data)
                                self.handle_element(element, data)
                            case "c_tag":
                                element = wait.until(EC.visibility_of_element_located((By.TAG_NAME, path)))
                                data = self.getCaptcha(data)
                                self.handle_element(element, data)
                            case "c_link_text":
                                element = wait.until(EC.visibility_of_element_located((By.LINK_TEXT, path)))
                                data = self.getCaptcha(data)
                                self.handle_element(element, data)
                            case "c_partial_link_text":
                                element = wait.until(EC.visibility_of_element_located((By.PARTIAL_LINK_TEXT, path)))
                                data = self.getCaptcha(data)
                                self.handle_element(element, data)
                            case "c_css_selector":
                                element = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, path)))
                                data = self.getCaptcha(data)
                                self.handle_element(element, data)
                            case "c_xpath":
                                element = wait.until(EC.visibility_of_element_located((By.XPATH, path)))
                                data = self.getCaptcha(data)
                                self.handle_element(element, data)
                            case "enter":
                                pyautogui.press('enter')
                            case "sleep":
                                time.sleep(data)
                        time.sleep(1)  # Wait for one second to observe the effect

                    except (NoSuchElementException, TimeoutException, NoSuchWindowException):
                        self.error_signal.emit(f"Element not found or timeout for Path: {path}")
                    except WebDriverException as e:
                        self.error_signal.emit(f"WebDriver error: {str(e)}")
        except WebDriverException as e:
            self.error_signal.emit(f"WebDriver initialization error: {str(e)}")
        finally:
            self.finished.emit()

    def handle_element(self, element, data):
        if element.tag_name in ["input", "textarea"] and data != "Enter Data":
            element.clear()
            element.send_keys(data)
        elif element.tag_name == "select":
            from selenium.webdriver.support.ui import Select
            select = Select(element)
            select.select_by_visible_text(data)
        else:
            element.click()

    def download_captcha(self, url, filepath):
        response = requests.get(url)
        with open(filepath, 'wb') as file:
            file.write(response.content)

    def recognize_captcha(self, filepath):
        ocr = ddddocr.DdddOcr()
        with open(filepath, 'rb') as file:
            image_bytes = file.read()
        result = ocr.classification(image_bytes)
        return result
    def getCaptcha(self, xpath):
        wait = WebDriverWait(self.driver, 30)
        element = wait.until(EC.visibility_of_element_located((By.XPATH, xpath)))
        captcha_url = element.get_attribute('src')
        self.download_captcha(captcha_url, 'captcha.jpg')
        data = self.recognize_captcha('captcha.jpg')
        return data

    def stop(self):
        self.running = False
        if self.driver:
            try:
                self.driver.quit()
            except WebDriverException:
                pass

class ComboBoxDelegate(QStyledItemDelegate):
    def createEditor(self, parent, option, index):
        combo_box = QComboBox(parent)
        combo_box.addItems(["url", "xpath", "iframein", "iframeout", "id", "name", "class_name", "tag", "link_text", "partial_link_text", "css_selector", "c_id", "c_xpath", "c_name", "c_class_name", "c_css_selector", "c_link_text", "c_partial_link_text", "c_tag"])
        return combo_box

    def setEditorData(self, editor, index):
        text = index.model().data(index, Qt.EditRole)
        editor.setCurrentText(text)

    def setModelData(self, editor, model, index):
        model.setData(index, editor.currentText(), Qt.EditRole)

class TextEditDelegate(QStyledItemDelegate):
    def createEditor(self, parent, option, index):
        text_edit = QTextEdit(parent)
        return text_edit

    def setEditorData(self, editor, index):
        text = index.model().data(index, Qt.EditRole)
        editor.setText(text)

    def setModelData(self, editor, model, index):
        model.setData(index, editor.toPlainText(), Qt.EditRole)

class SeleniumTool(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.driver = None
        self.worker = None

    def initUI(self):
        self.setWindowTitle('Selenium Tool')
        self.setGeometry(100, 100, 800, 600)

        main_widget = QWidget(self)
        self.setCentralWidget(main_widget)

        main_layout = QHBoxLayout(main_widget)

        self.left_widget = QWidget(self)
        self.left_layout = QVBoxLayout(self.left_widget)
        main_layout.addWidget(self.left_widget, 8)

        self.right_widget = QWidget(self)
        self.right_layout = QVBoxLayout(self.right_widget)
        main_layout.addWidget(self.right_widget, 2)

        self.url_edit = QLineEdit(self)
        self.url_edit.setPlaceholderText("Enter URL")
        self.left_layout.addWidget(self.url_edit)

        self.tree = QTreeView(self)
        self.model = QStandardItemModel()
        self.model.setHorizontalHeaderLabels(['Mode', 'Path', 'Data'])
        self.tree.setModel(self.model)
        self.left_layout.addWidget(self.tree)

        self.tree.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tree.customContextMenuRequested.connect(self.show_context_menu)

        self.add_button = QPushButton('Add', self)
        self.add_button.clicked.connect(self.add_row)
        self.left_layout.addWidget(self.add_button)

        self.start_button = QPushButton('Start', self)
        self.start_button.clicked.connect(self.start_execution)
        self.right_layout.addWidget(self.start_button)

        self.stop_button = QPushButton('Stop', self)
        self.stop_button.clicked.connect(self.stop_execution)
        self.right_layout.addWidget(self.stop_button)

        self.export_button = QPushButton('Export', self)
        self.export_button.clicked.connect(self.export_data)
        self.right_layout.addWidget(self.export_button)

        self.import_button = QPushButton('Import', self)
        self.import_button.clicked.connect(self.import_data)
        self.right_layout.addWidget(self.import_button)

        # Set delegates for the columns
        self.tree.setItemDelegateForColumn(0, ComboBoxDelegate(self.tree))
        self.tree.setItemDelegateForColumn(2, TextEditDelegate(self.tree))

    def add_row(self):
        mode_item = QStandardItem("xpath")
        mode_item.setEditable(True)
        path_item = QStandardItem("Enter Path")
        data_item = QStandardItem("Enter Data")
        self.model.appendRow([mode_item, path_item, data_item])

    def show_context_menu(self, position):
        menu = QMenu()
        delete_action = menu.addAction("Delete")
        delete_action.triggered.connect(self.delete_row)
        menu.exec_(self.tree.viewport().mapToGlobal(position))

    def delete_row(self):
        indexes = self.tree.selectionModel().selectedRows()
        if indexes:
            index = indexes[0]
            self.model.removeRow(index.row())

    def start_execution(self):
        url = self.url_edit.text().strip()
        if not url:
            QMessageBox.critical(self, "Error", "Please enter a valid URL")
            return

        if self.worker:
            self.worker.stop()
            self.worker.wait()

        if self.driver is not None:
            try:
                self.driver.title  # Try to access the driver to check if it's still valid
            except WebDriverException:
                self.driver = None

        self.worker = SeleniumWorker(url, self.model, driver=self.driver)
        self.worker.finished.connect(self.on_worker_finished)
        self.worker.driver_created.connect(self.set_driver)  # Connect to the signal to get the driver instance
        self.worker.error_signal.connect(self.show_error)  # Connect error signal to show_error slot
        self.worker.start()

    def stop_execution(self):
        if self.worker:
            self.worker.stop()
            self.worker.wait()

    def set_driver(self, driver):
        self.driver = driver

    @pyqtSlot(str)
    def show_error(self, message):
        QMessageBox.critical(self, "Error", message)

    def on_worker_finished(self):
        self.worker = None

    def export_data(self):
        url = self.url_edit.text().strip()
        if not url:
            QMessageBox.critical(self, "Error", "URL field cannot be empty")
            return

        data = {
            'URL': [url] * self.model.rowCount(),
            'Mode': [self.model.item(row, 0).text() for row in range(self.model.rowCount())],
            'Path': [self.model.item(row, 1).text() for row in range(self.model.rowCount())],
            'Data': [self.model.item(row, 2).text() for row in range(self.model.rowCount())],
        }

        df = pd.DataFrame(data)

        file_name, ok = QInputDialog.getText(self, "Input File Name", "Enter the file name:")
        try:
            if ok and file_name:
                if not os.path.exists('config'):
                    os.makedirs('config')
                file_path = os.path.join('config', f"{file_name}.xlsx")
                df.to_excel(file_path, index=False)
                QMessageBox.information(self, "Success", f"Data exported successfully to {file_path}")
        except:
            pass

    def import_data(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Open Config File", "", "Excel Files (*.xlsx)")
        if not file_path:
            return

        df = pd.read_excel(file_path)
        self.url_edit.setText(df.at[0, 'URL'])
        self.model.removeRows(0, self.model.rowCount())
        for index, row in df.iterrows():
            mode_item = QStandardItem(str(row['Mode']))
            path_item = QStandardItem(row['Path'])
            data_item = QStandardItem(row['Data'])
            self.model.appendRow([mode_item, path_item, data_item])
        QMessageBox.information(self, "Success", "Data imported successfully")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = SeleniumTool()
    ex.show()
    sys.exit(app.exec_())
