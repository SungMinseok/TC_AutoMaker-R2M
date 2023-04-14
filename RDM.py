
import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QLabel, QPushButton, QFileDialog, QTextEdit, QComboBox
from enum import Enum, auto
import datetime
import os
import CLMaker_cashshop as ClCash
import CLMaker_Event as ClEvent

class DocumentType(Enum):
    TestCase = auto()
    CheckList = auto()

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("R2M Maker 0.1")
        self.setGeometry(100, 100, 300, 300)

        self.contents_label = QLabel("컨텐츠", self)
        self.contents_label.setGeometry(20, 20, 150, 30)
        self.contents_label.move(20, 20)

        self.contents_combo = QComboBox(self)
        self.contents_combo.move(180, 20)
        self.contents_combo.addItem("유료상점")
        self.contents_combo.addItem("이벤트")

        # Create a label for the file type input
        self.file_type_label = QLabel("문서 타입", self)
        self.file_type_label.setGeometry(20, 220, 150, 30)
        self.file_type_label.move(20, 70)

        # # Create a text edit for the file type input
        # self.file_type_input = QTextEdit(self)
        # self.file_type_input.setGeometry(200, 20, 150, 30)
                # Create a ComboBox widget for the file type input
        self.file_type_combo = QComboBox(self)
        self.file_type_combo.move(180, 70)
        self.file_type_combo.addItem("TC")
        self.file_type_combo.addItem("CL")


        # Create a button for selecting the data file
        self.data_file_button = QPushButton("데이터 파일 선택", self)
        self.data_file_button.setGeometry(20, 120, 150, 30)
        self.data_file_button.clicked.connect(self.select_data_file)

        # Create a label for displaying the selected data file
        self.data_file_label = QLabel("", self)
        self.data_file_label.setGeometry(200, 120, 150, 30)

        # Create a button for starting the processing
        self.process_button = QPushButton("처리 시작", self)
        self.process_button.setGeometry(80, 170, 150, 30)
        self.process_button.clicked.connect(self.start_processing)

        # Create a label for displaying the processing status
        self.status_label = QLabel("", self)
        self.status_label.setGeometry(20, 170, 200, 30)

    def select_data_file(self):
        # Open a file dialog to select the data file
        data_file, _ = QFileDialog.getOpenFileName(self, "데이터 파일 선택")
        self.data_file_label.setText(data_file)

    def start_processing(self):
        # Get the selected file type and data file
        file_type_input = self.file_type_input.toPlainText()
        data_file = self.data_file_label.text()

        # Validate the file type input
        try:
            file_type = DocumentType(int(file_type_input))
        except ValueError:
            self.status_label.setText("잘못된 입력입니다.")
            return

        # Validate the data file input
        if not os.path.isfile(data_file):
            self.status_label.setText("데이터 파일을 선택하세요.")
            return

        # Process the data file
        todayDate = datetime.datetime.today().date()
        #if file_type == DocumentType.TestCase:
        global dateID
        dateID= 0
        dateID = (3,"목")

        days_until_target = (dateID[0] - todayDate.weekday()) % 7
        thursdayDate = todayDate + datetime.timedelta(days=days_until_target)
        global check_start_date
        check_start_date = thursdayDate.strftime('%Y-%m-%d')

        print(f"이번주 {dateID[1]}요일 {check_start_date} 기준으로 작성됩니다.")
        check_start_date = datetime.datetime.strptime(check_start_date, '%Y-%m-%d')

        # check_start_date = input("업데이트날짜(YYYY-MM-DD) >: ")
        # if check_start_date == "":
        #     check_start_date = datetime.datetime.now().strftime('%Y-%m-%d')
        result_directory = f"./Event_{file_type}"
        if not os.path.isdir(result_directory):
            os.mkdir(result_directory)

        global xl_filename
        xl_filename = f"{result_directory}/result_{time.strftime('%y%m%d_%H%M%S')}.xlsx"
        targetList = extract_data(data_file)
        write_data(targetList)
        postprocess_cashshop()
        self.status_label.setText("처리가 완료되었습니다.")
        #else:
        #    self.status_label.setText("지원하지 않는 문서 타입입니다.")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()
    sys.exit(app.exec_())
