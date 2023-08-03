import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QLabel
from PyQt5.QtCore import QTimer

class TqdmProgressApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()

        self.progressLabel = QLabel(self)
        layout.addWidget(self.progressLabel)

        self.setLayout(layout)

        self.setGeometry(100, 100, 300, 100)
        self.setWindowTitle('Animated Progress in PyQt5')

        self.start()

    def start(self):
    

        # Start the timer to update the progress label every 200 milliseconds (adjust as needed).
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_progress_label)
        self.timer.start(100)

        self.log_texts = ["□", "■"]
        self.log_index = 0

    def update_progress_label(self):
        log_text = self.log_texts[self.log_index]
        self.progressLabel.setText(log_text)

        # Rotate the log_texts list to show the next text in the next iteration.
        self.log_index = (self.log_index + 1) % len(self.log_texts)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = TqdmProgressApp()
    window.show()
    sys.exit(app.exec_())
