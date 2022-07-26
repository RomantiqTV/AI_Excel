import pandas
import selenium
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
from selenium.webdriver.common.by import By
import time
from urllib.request import Request, urlopen
from selenium.webdriver.chrome.options import Options
import pyperclip as pc
import requests
import wget
import zipfile
import os
import re
from webdriver_auto_update import check_driver
import tempfile
import contextlib
import sys
import readchar
from PyQt5.QtWidgets import QWidget, QMainWindow, QApplication, QFrame, QLabel, QPushButton, QLineEdit
from PyQt5 import QtCore, QtGui, QtWidgets, uic
from PyQt5.QtCore import *
from PyQt5.QtGui import QMovie
from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5 import QtTest
import sys


class UI(QMainWindow):
    def __init__(self):
        super(UI, self).__init__()
        # Load UI file
        uic.loadUi("new_excel_gui.ui", self)

        self.setWindowFlag(Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)

        #Define widgets
        self.widget = self.findChild(QWidget, name="centralwidget")
        self.label_1 = self.findChild(QLabel, name="label_4")
        self.label_2 = self.findChild(QLabel, name="label_5")
        self.generate_text = self.findChild(QLabel, name="label_6")
        self.generate_text.hide()
        self.copy_button = self.findChild(QPushButton, "pushButton_6")
        self.copy_button.hide()
        self.load_label = self.findChild(QLabel, name="label")
        self.load_label.setScaledContents(True)
        self.load_label.hide()
        self.lineedit = self.findChild(QLineEdit, name="lineEdit_2")
        self.exit_button = self.findChild(QPushButton, name="pushButton_3")
        self.cancel_button = self.findChild(QPushButton, name="pushButton_5")
        self.cancel_button.hide()
        self.new_button = self.findChild(QPushButton, name="pushButton_7")
        self.new_button.hide()
        self.generate_button = self.findChild(QPushButton, name="pushButton_4")

        self.exit_button.clicked.connect(self.close)
        self.generate_button.clicked.connect(self.GeneratorClick)
        self.copy_button.clicked.connect(self.CopyButton)

        self.movie = QMovie("loading.gif")
        self.load_label.setMovie(self.movie)

        self.show()

    def startAnimation(self):
        self.movie.start()

        # Stop Animation(According to need)
    def stopAnimation(self):
        self.movie.stop()

    def mousePressEvent(self, event):
        self.oldPosition = event.globalPos()

        #Action2
    def mouseMoveEvent(self, event):
        delta = QPoint(event.globalPos() - self.oldPosition)
        self.move(self.x() + delta.x(), self.y() + delta.y())
        self.oldPosition = event.globalPos()

    def GeneratorClick(self):
        global UserInput
        UserInput = self.lineedit.text()
        self.lineedit.setDisabled(True)
        self.copy_button.hide
        self.lineedit.setReadOnly(True)
        self.load_label.show()
        self.generate_text.show()
        self.generate_button.hide()
        self.startAnimation()

        self.worker = Worker_thread()
        self.worker.start()
        self.worker.update_response.connect(self.evt_update_response)
        self.worker.finished.connect(self.on_finished)

    def on_finished(self):
        self.lineedit.setDisabled(False)
        self.lineedit.setStyleSheet("QLineEdit {font-weight: bold;}"
                                    )
        self.stopAnimation()
        self.load_label.hide()
        self.generate_text.hide()
        self.copy_button.show()
        self.label_2.setText("AI Excel formula output:")
        self.new_button.show()
        self.new_button.clicked.connect(self.NewButton)

    def NewButton(self):
        self.lineedit.setStyleSheet("")
        self.lineedit.setDisabled(False)
        self.lineedit.setReadOnly(False)
        self.generate_button.show()
        self.copy_button.hide()
        self.new_button.hide()
        self.lineedit.setText("")

    def CopyButton(self):
        self.lineedit.selectAll()
        self.lineedit.copy()
        self.label_5.setText("Formula has been copied")


    def evt_update_response(self, val):
        QtTest.QTest.qWait(1500)
        self.lineedit.setText(val)

class Worker_thread(QThread):
    update_response = pyqtSignal(str)

    def run(self):
        self.threadactive = True
        check_driver(tempfile.gettempdir())

        s = Service(tempfile.gettempdir() + '//chromedriver.exe')
        chrome_options = Options()
        chrome_options.add_argument("--headless")
        driver = webdriver.Chrome(service=s, options=chrome_options)
        driver.get("https://excelformulabot.com/")
        time.sleep(2)
        SearchBar = driver.find_element(By.XPATH, "/html/body/div[1]/div[4]/div/textarea")
        SearchBar.send_keys(UserInput)
        time.sleep(1)
        QueryEnter = driver.find_element(By.XPATH, "/html/body/div[1]/div[4]/div/button").click()
        time.sleep(4)
        QueryResponse = driver.find_element(By.XPATH, "/html/body/div[1]/div[4]/div/div[6]/div[2]").text
        pattern = r'^[^=]{1}'
        if re.search(pattern, QueryResponse):
            sep = '='
            QueryResponse = QueryResponse.split(sep, 1)[1]
            QueryResponse = "=" + QueryResponse
            print(QueryResponse)
        else:
            QueryResponse = QueryResponse
        self.update_response.emit(QueryResponse)
        driver.close()
        driver.quit()

class Worker_thread2(QThread):
    def run(self):
        s = Service(tempfile.gettempdir() + '//chromedriver.exe')
        chrome_options = Options()
        chrome_options.add_argument("--headless")
        driver = webdriver.Chrome(service=s, options=chrome_options)
        driver.close()
        driver.quit()

#Initialize the app
app = QApplication(sys.argv)
UIWindow = UI()
app.exec()
