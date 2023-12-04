# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file '230509.ui'
#
# Created by: PyQt5 UI code generator 5.15.7
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.

import sys
from PyQt5.QtWidgets import *
from PyQt5 import uic
from PyQt5.QtCore import Qt
from PyQt5 import QtCore, QtWidgets
from PyQt5.QtGui import QKeySequence,QPixmap, QColor
from PyQt5.QtWidgets import QLabel, QApplication, QWidget, QVBoxLayout
from PyQt5.QtCore import QDate,QTimer,Qt, QThread, pyqtSignal
import time
import threading
import traceback
#import socket
import os
import pandas as pd
import mspkg as mp



form_class = uic.loadUiType(f'RDM_UI.ui')[0]
FROM_CLASS_Loading = uic.loadUiType("load.ui")[0]
#화면을 띄우는데 사용되는 Class 선언            print(e)

user_name = os.getlogin()
patch_file_name = 'release_note_RDM.xlsx'
patch_show_count = 5


cache_folder = "./cache"
if not os.path.isdir(cache_folder):                                                           
    os.mkdir(cache_folder)

cache_path = f'./cache/cache_{user_name}.csv'
try:
    df_cache = pd.read_csv(cache_path, sep='\t', encoding='utf-16', index_col='key')
except FileNotFoundError as e:
    print(f'{e} : 캐시 파일 없음')

class WindowClass(QMainWindow, form_class) :
    def __init__(self) :
        super().__init__()
        self.setupUi(self) 
        self.setGeometry(1470,28,400,400)
        self.setFixedSize(450,350)
        self.action_patchnote.triggered.connect(lambda : self.파일열기(patch_file_name))
        self.print_log("실행 가능")

        self.worker_thread = None
        #self.worker_thread.update_message.connect(self.display_error_message)
#복붙시작◈◈◈◈◈◈◈◈◈◈◈◈◈◈◈◈◈◈◈◈◈◈◈◈◈◈◈◈◈◈◈◈◈◈◈◈◈◈◈◈◈◈◈◈◈◈◈◈◈◈◈◈◈◈◈◈◈◈◈◈◈◈◈◈
        self.set_data_path()
        self.set_result_path()


        file_dict = mp.get_recent_file_list(os.getcwd())
        last_modified_date = list(file_dict.values())[0]

        self.btn_datapath.clicked.connect(self.select_data_file)
        self.btn_datapath_2.clicked.connect(lambda : self.파일열기(self.input_datapath.text()))
        self.btn_execute.clicked.connect(self.activate)

        self.input_datapath.setAcceptDrops(True)
        self.input_datapath.dragEnterEvent = self.drag_enter_event
        self.input_datapath.dropEvent = self.drop_event

        self.combox_country.currentTextChanged.connect(self.set_data_path)
        self.combox_contents.currentTextChanged.connect(self.set_data_path)

        self.combox_contents.currentTextChanged.connect(self.set_result_path)
        self.combox_doctype.currentTextChanged.connect(self.set_result_path)

        self.dateedit.setDate(QDate.currentDate())

        self.btn_resultpath.clicked.connect(self.select_data_file)
        self.btn_resultpath_2.clicked.connect(lambda : self.파일열기(self.input_resultpath.text()))    

        

        '''패치노트'''
        patch_note_check = self.import_cache_all([QCheckBox,'checkBox_99'])
        if patch_note_check != None and patch_note_check.lower() == 'true': #or ( patch_note_check.lower() == 'false' and is_next_day): 
            x, patch_see_again = self.popup2(des_text=
                                        f"업데이트 일자 : {last_modified_date}\n\n최신 업데이트 항목 {patch_show_count}개\n\n{mp.read_patch_notes(patch_file_name,patch_show_count)}", popup_type='patchnote')
        
            self.checkBox_99.setChecked(not patch_see_again)

        self.import_cache_all()

#기본동작■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■

    def select_data_file(self):
        # Open a file dialog to select the data file
        file_filter = "Video files (*.mp4 *.mkv)"
        file_filter = "엑셀 파일 (*.xlsx)"

        data_file, _ = QFileDialog.getOpenFileName(self,"데이터 파일 선택", filter= file_filter)
        self.input_datapath.setText(data_file)

    def drag_enter_event(self, event):
        # 드래그앤드랍 가능한 MIME 타입인지 체크
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()
    
    def drop_event(self, event):
        # 파일 경로를 가져와서 QLineEdit에 입력
        urls = event.mimeData().urls()
        file_path = urls[0].toLocalFile()
        self.input_datapath.setText(file_path)

    def set_data_path(self):
        contents = self.combox_contents.currentText()
        country = self.combox_country.currentText()

        self.input_datapath.setText(f"{contents}DATA_{country}.xlsx")

    def set_result_path(self):
        contents = self.combox_contents.currentText()
        document_type = self.combox_doctype.currentText()
        
        result_path = f"{contents}_{document_type}"

        self.input_resultpath.setText(result_path)
            
        if not os.path.isdir(result_path) :
            os.mkdir(result_path)

    def 파일열기(self,filePath):
        try:
            os.startfile(filePath)
        except : 
            print("파일 없음 : "+filePath)    

    def print_log(self, log): # / - \ / - \ / ㅡ ㄷ
        self.progressLabel.setText(log)
        QApplication.processEvents()

#■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
    def make_process(self, result_path, data_file_name, contents_name, doctype, date_text, check_box_list):
        
        
        try:
            #self.popUp("454545")
            if contents_name == "유료상점" : 
                if doctype == "CheckList" :
                    data = ClCash.extract_data_cashshop(data_file_name,date_text)
                    result_file_name = ClCash.write_data_cashshop_inspection(data,result_path,check_box_list)
                    #ClCash.postprocess_cashshop(result_file_name)
                elif doctype == "TestCase" :
                    data = ClCash.extract_data_cashshop(data_file_name, date_text)
                    result_file_name = ClCash.write_data_cashshop(data,result_path)

                
                ClCash.postprocess_cashshop(result_file_name)

            elif contents_name == "이벤트" :
                if doctype == "CheckList" :
                    data = ClEvent.extract_data(data_file_name, date_text)
                    result_file_name = ClEvent.write_data(data,result_path,check_box_list)
                    #ClEvent.postprocess_cashshop(result_file_name)
                elif doctype == "TestCase" :
                    data = ClEvent.extract_data(data_file_name,date_text)
                    result_file_name = ClEvent.write_data_event_testcase(data,result_path)

                if len(data) == 0:
                    
                    return

                ClEvent.postprocess_cashshop(result_file_name)

            os.startfile(os.path.normpath(result_file_name))

        except Exception as e:
            #log_file = fr".\log\log_error_{user_name}.txt"
            msg = f'{contents_name=}\n{doctype=}\n{date_text=}\n'
            msg += traceback.format_exc()
            self.make_log(msg,auto_open=True)
            # with open(log_file, "a") as file:
            #     file.write(f'{time.strftime("%y%m%d_%H%M%S")}\n{error_message}\n')
            # print(f'생성실패 : {e}')
            # os.startfile(log_file)

        self.worker_thread.finished.emit()
        self.worker_thread.quit()
        #self.worker_thread.wait()



        self.loading
        self.loading.deleteLater()

    #result_path, data_file_name
        msg = f'{result_path=}\n{data_file_name=}\n{contents_name=}\n{doctype=}\n{date_text=}\n'
        self.make_log(msg,'execute')

    def make_log(self, msg, log_type : str = 'error', auto_open :bool = False):
        '''
        log_type : str = error / execute
        '''


        
        log_file = fr".\log\log_{log_type}.txt"
        #error_message = traceback.format_exc()
        if not os.path.exists(log_file):
            with open(log_file, "w"):
                pass  # Create the file if it doesn't exist
        with open(log_file, "r") as file:        
            existing_logs = file.read()


        with open(log_file, "w") as file:
            file.write(f'\n\
user={user_name}\n\
date={time.strftime("%Y-%m-%d %H:%M:%S")}\n\
{msg}\n\
────────────────────────────────────────\n\
{existing_logs}\
    ')
        #print(f'생성실패 : {e}')
        if auto_open :
            os.startfile(log_file)

    def activate(self):

        
        result_path = self.input_resultpath.text()
        data_file_name = self.input_datapath.text()
        contents_name = self.combox_contents.currentText()#유료상점/이벤트
        doctype = self.combox_doctype.currentText()#체크리스트/테스트케이스
        date_text = self.dateedit.text()

        check_box_list = [
            self.checkBox_0.isChecked(),
            self.checkBox_1.isChecked(),
            self.checkBox_2.isChecked(),
            self.checkBox_3.isChecked(),
            ]
        #self.start_loading()

        # loading_thread = threading.Thread(target = self.start_loading(self))
        # loading_thread.start()
        self.loading = loading(self)
        #time.sleep(1)
        # self.worker = Worker(self)
        # self.worker.start()
        # loading_window = LoadingWindow()
        # loading_thread = LoadingThread()

        # loading_thread.finished.connect(loading_window.close)

        # loading_window.show()
        # loading_thread.start()
        #self.show_loading_window()    
    
        self.worker_thread = WorkerThread(myWindow,result_path, data_file_name,contents_name, doctype, date_text,check_box_list)
        
        self.worker_thread.finished.connect(self.cleanup)
        self.worker_thread.start()
        #self.make_process(result_path,data_file_name,contents_name,doctype,date_text,check_box_list)

        #making_thread = threading.Thread(target = self.make_process(result_path,data_file_name,contents_name,doctype,date_text,check_box_list))
        #making_thread.start()
        #self.loading = loading(self)

        
        #self.loading
        #self.loading.deleteLater()
        # result_file_name = ""
        # if self.combox_contents.currentText() == "유료상점" :
        #     if self.combox_doctype.currentText() == "CheckList" :
        #         self.print_log("데이터 추출 중...")
        #         data = ClCash.extract_data_cashshop(data_file_name, self.dateedit.text())
        #         self.print_log("데이터 쓰는 중...")
        #         result_file_name = ClCash.write_data_cashshop_inspection(data,result_path,check_box_list)
        #         self.print_log("데이터 정리 중...")
        #         ClCash.postprocess_cashshop(result_file_name)
        #     elif self.combox_doctype.currentText() == "TestCase" :
        #         self.print_log("데이터 추출 중...")
        #         data = ClCash.extract_data_cashshop(data_file_name, self.dateedit.text())
        #         self.print_log("데이터 쓰는 중...")
        #         result_file_name = ClCash.write_data_cashshop(data,result_path)
        #         self.print_log("데이터 정리 중...")
        #         ClCash.postprocess_cashshop(result_file_name)


        # elif self.combox_contents.currentText() == "이벤트" :
        #     if self.combox_doctype.currentText() == "CheckList" :
        #         self.print_log("데이터 추출 중...")
        #         data = ClEvent.extract_data(data_file_name, self.dateedit.text())
        #         self.print_log("데이터 쓰는 중...")
        #         result_file_name = ClEvent.write_data(data,result_path)
        #         self.print_log("데이터 정리 중...")
        #         ClEvent.postprocess_cashshop(result_file_name)
        #     elif self.combox_doctype.currentText() == "TestCase" :
        #         self.print_log("데이터 추출 중...")
        #         data = ClEvent.extract_data(data_file_name, self.dateedit.text())
        #         self.print_log("데이터 쓰는 중...")
        #         result_file_name = ClEvent.write_data_event_testcase(data,result_path)
        #         self.print_log("데이터 정리 중...")
        #         ClEvent.postprocess_cashshop(result_file_name)

    

        # self.print_log("문서 여는 중...")
        #os.startfile(os.path.normpath(result_file_name))
        # self.print_log("실행 가능")
    def cleanup(self):
        self.worker_thread = None

    def start_loading(self,qma):
        loading(self)
        # loading_thread = threading.Thread(target = loading(self))
        # loading_thread.start()
    def popUp(self,desText,titleText="error"):
        #if type == "about" :
        #print('345435')
        msg = QtWidgets.QMessageBox()  
        msg.setGeometry(1520,28,400,2000)
        msg.setText(desText)

        #msg.setTextInteractionFlags(QtCore.Qt.TextInteractionFlag.TextSelectableByMouse)
        msg.setTextInteractionFlags(QtCore.Qt.TextInteractionFlag.TextSelectableByMouse)
        msg.addButton(QtWidgets.QMessageBox.Ok)
        result = msg.exec_()
        
    def display_error_message(self, error_message):
        self.popUp(desText=error_message)
        print(f'생성실패 : {error_message}')
        os.system('pause')

    def import_cache_all(self,any_widget = None):
        '''any_widget : [QLineEdit,'input_00']'''

        try:
            if df_cache is None :
            # Load CSV file with tab delimiter and utf-16 encoding
                df = pd.read_csv(cache_path, sep='\t', encoding='utf-16', index_col='key')
            else :
                df = df_cache
            if any_widget == None :
                all_widgets = self.findChildren((QLineEdit, QLabel, QComboBox, QCheckBox, QPlainTextEdit,QPushButton))
            else:
                all_widgets = [self.findChild(any_widget[0] ,any_widget[1])]
                #return 
            #all_widgets = self.findChildren((QLineEdit, QLabel, QComboBox, QCheckBox, QPlainTextEdit,QPushButton))

            for widget in all_widgets:
                object_name = widget.objectName()
                if object_name in df.index:
                    value = str(df.loc[object_name, 'value'])
                    if isinstance(widget, (QLineEdit,QLabel,QPushButton)):
                        widget.setText(value)
                    elif isinstance(widget, QComboBox):
                        # Set selected index based on the value, adjust as needed
                        index = widget.findText(value)
                        if index != -1:
                            widget.setCurrentIndex(index)
                    elif isinstance(widget, QCheckBox):
                        widget.setChecked(value.lower() == 'true')
                    elif isinstance(widget, QPlainTextEdit):
                        widget.setPlainText(value)

            if any_widget != None :
                return value
        except Exception as e:
            print(f"Error importing cache: {e}")



    def export_cache_all(self):
        try:
            data = {'key': [], 'value': []}

            all_widgets = self.findChildren((QLineEdit, QLabel, QComboBox, QCheckBox, QPlainTextEdit,QPushButton, QDateTimeEdit))

            for widget in all_widgets:
                value = ""
                if isinstance(widget, (QLineEdit,QLabel)) :
                    value = widget.text()
                elif isinstance(widget, (QPushButton)) :
                    if 'preset_bookmark' in widget.objectName() : 
                        value = widget.text()
                    else : 
                        continue
                elif isinstance(widget, QComboBox):
                    value = widget.currentText()
                elif isinstance(widget, QCheckBox):
                    value = str(widget.isChecked())
                elif isinstance(widget, QPlainTextEdit):
                    value = widget.toPlainText()
                elif isinstance(widget, QDateTimeEdit):
                    value = widget.dateTime().toString(Qt.ISODate)

                if value != "":
                    key = widget.objectName()
                    data['key'].append(key)
                    data['value'].append(value)

            df = pd.DataFrame(data)
            df.set_index('key', inplace=True)
            df.to_csv(cache_path, sep='\t', encoding='utf-16')
            print(f"exporting cache successfully!")
        except Exception as e:
            print(f"Error exporting cache: {e}")
    
    def popup2(self, des_text = "", popup_type = ''):

        msg = QtWidgets.QMessageBox()  
        #msg.setGeometry(1470,58,300,2000)
        msg.setTextInteractionFlags(QtCore.Qt.TextInteractionFlag.TextSelectableByMouse)
        
        if popup_type == "":
            msg.setStandardButtons(QtWidgets.QMessageBox.Ok | QtWidgets.QMessageBox.Cancel)
            msg.setText(des_text)
            return msg.exec_()
        elif popup_type == 'patchnote' :
                        # Create a checkbox
            checkbox = QtWidgets.QCheckBox("오늘은 그만 보기", msg)
            #msg.setStandardButtons(QtWidgets.QMessageBox.Open | QtWidgets.QMessageBox.Cancel)
            msg.setStandardButtons(QtWidgets.QMessageBox.Cancel)

            msg.setCheckBox(checkbox)
            msg.setText(des_text)
            #print(checkbox.isChecked())
            return msg.exec_(), checkbox.isChecked()
        elif popup_type == 'report' :
            #checkbox = QtWidgets.QCheckBox("오늘은 그만 보기", msg)
            #msg.setStandardButtons(QtWidgets.QMessageBox.Open | QtWidgets.QMessageBox.Cancel)
            report_text = QtWidgets.QPlainTextEdit()

            msg.setStandardButtons(QtWidgets.QMessageBox.Ok | QtWidgets.QMessageBox.Cancel)

            #msg.setText(des_text)
            #print(checkbox.isChecked())
            return msg.exec_(), checkbox.isChecked()

    def closeEvent(self,event):
        self.export_cache_all()

class loading(QWidget,FROM_CLASS_Loading):
    
    def __init__(self,parent):
        super(loading, self).__init__(parent)    
        self.setupUi(self) 
        #self.resize(parent.size())
        self.setFixedSize(parent.size())
        self.center()
        # Get the size of the parent widget and set the loading widget to the same size
        
        self.show()
        
        #self.movie = QMovie('rengar.gif', QByteArray(), self)
        self.movie = QMovie('dda956507874240e5f0d05ac575e1c30.webp', QByteArray(), self)
        self.movie.setCacheMode(QMovie.CacheAll)
        self.label.setMovie(self.movie)
        self.label.setScaledContents(True)
        #self.movie.set(500,500)
        self.movie.start()
        self.setWindowFlags(Qt.FramelessWindowHint)
    # 위젯 정중앙 위치
    def center(self):
        
        size=self.size()
        ph = self.parent().geometry().height()
        pw = self.parent().geometry().width()
        self.move(int(pw/2 - size.width()/2), int(ph/2 - size.height()/2))
        self.move(int(pw/2 - size.width()/2), int(ph/2 - size.height()/2))
class WorkerThread(QThread):
    finished = pyqtSignal()
    #update_message = pyqtSignal(str)  # Define a custom signal
    def __init__(self, window, result_path, data_file_name, contents_name, doctype, date_text, check_box_list):
        super().__init__()
        self.window = window
        self.result_path = result_path
        self.data_file_name = data_file_name
        self.contents_name = contents_name
        self.doctype = doctype
        self.date_text = date_text
        self.check_box_list = check_box_list

    def run(self):
        #try:
        self.window.make_process(self.result_path, self.data_file_name, self.contents_name, self.doctype, self.date_text, self.check_box_list)

            #os.system('pause')
        self.finished.emit()



from PyQt5 import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from PyQt5.QtWidgets import*
from PyQt5 import uic
import sys
from enum import Enum, auto
import datetime
import os
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QApplication, QMainWindow, QLabel, QPushButton, QFileDialog, QTextEdit, QComboBox

import CLMaker_cashshop as ClCash
import CLMaker_Event as ClEvent
import openpyxl

# .py 파일을 모두 찾아내는 함수
def get_recent_file_list(directory):
    '''
    os.getcwd()

    
    #print(list(file_dict.keys())[0])
    #print(list(file_dict.values())[0])
    '''
    
    current_directory = directory

    def find_py_files(directory):
        for root, dirs, files in os.walk(directory):
            for file in files:
                if file.endswith('.py'):
                    yield os.path.join(root, file)

    # .py 파일 중에서 수정 날짜가 최신인 순으로 정렬
    py_files = list(find_py_files(current_directory))
    py_files.sort(key=lambda x: os.path.getmtime(x), reverse=True)

    file_dict = {}

    # 최신 수정 날짜를 가진 .py 파일 출력
    for file in py_files:
        modified_time = os.path.getmtime(file)
        modified_date = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(modified_time))
        #print(f"{file}: {modified_date}")
        file_dict[file] = modified_date

    #print(list(file_dict.keys())[0])
    #print(list(file_dict.values())[0])
    return file_dict
if __name__ == "__main__" :
    #QApplication : 프로그램을 실행시켜주는 클래스
    app = QApplication(sys.argv) 

    #WindowClass의 인스턴스 생성
    myWindow = WindowClass() 

    #프로그램 화면을 보여주는 코드
    myWindow.show()

    #프로그램을 이벤트루프로 진입시키는(프로그램을 작동시키는) 코드
    app.exec_()