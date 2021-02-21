import sys
from PySide6.QtWidgets import (
    QLabel,
    QLineEdit,
    QPlainTextEdit,
    QPushButton,
    QApplication,
    QTextEdit,
    QVBoxLayout,
    QDialog,
)
import bible_class


class Form(QDialog):
    def __init__(self, parent=None):
        super(Form, self).__init__(parent)
        # Create widgets
        self.input1 = QLineEdit()
        self.input2 = QLineEdit()
        self.input3 = QPlainTextEdit()
        self.input4 = QLineEdit()
        self.button0 = QPushButton("소스 ppt 열기")
        self.button1 = QPushButton("교독문 입력")
        self.button2 = QPushButton("첫번째 찬송가 입력")
        self.button3 = QPushButton("본문 입력")
        self.button4 = QPushButton("두번째 찬송가 입력")
        self.button5 = QPushButton("전환애니메이션 추가")
        # Create layout and add widgets
        layout = QVBoxLayout()
        layout.addWidget(self.button0)
        layout.addWidget(QLabel("교독문 번호를 입력하세요"))
        layout.addWidget(self.input1)
        layout.addWidget(self.button1)
        layout.addWidget(QLabel("첫번째 찬송가 번호를 입력하세요"))
        layout.addWidget(self.input2)
        layout.addWidget(self.button2)
        layout.addWidget(self.input3)
        self.input3.setPlaceholderText("설교 한글 파일 내용을 복사(Ctrl+C) 붙여넣기(Ctrl+V) 하세요.")
        layout.addWidget(self.button3)
        layout.addWidget(QLabel("두번째 찬송가 번호를 입력하세요"))
        layout.addWidget(self.input4)
        layout.addWidget(self.button4)
        layout.addWidget(self.button5)

        # Set dialog layout
        self.setLayout(layout)
        # Add button signal to greetings slot
        self.button0.clicked.connect(self.script0)
        self.button1.clicked.connect(self.script1)
        self.button2.clicked.connect(self.script2)
        self.button3.clicked.connect(self.script3)
        self.button4.clicked.connect(self.script4)
        self.button4.clicked.connect(self.script5)

    def script0(self):
        path = "c:\\Users\\Administrator\\Desktop\\WorkSpace\\pyPptx\\오전예배 (16x9)_2021_base.pptx"
        self.src_ppt = bible_class.Powerpoint()
        self.src_ppt.init_app()
        src_prs = self.src_ppt.open_prs(path=path)

    def script1(self):
        re_no = self.input1.text()
        bible_class.responsive_reading_copy(
            re_no, src_ppt=self.src_ppt, section_index=3
        )

    def script2(self):
        hymn_number = self.input2.text()
        bible_class.hymn_copy(hymn_number, src_ppt=self.src_ppt, section_index=4)

    def script3(self):
        raw = self.input3.toPlainText()
        bible_class.text_copy(raw, src_ppt=self.src_ppt)
        pass

    def script4(self):
        hymn_number = self.input4.text()
        bible_class.hymn_copy(hymn_number, src_ppt=self.src_ppt, section_index=8)
        pass

    def script5(self):
        self.src_ppt.change_transition()
        pass


if __name__ == "__main__":
    # Create the Qt Application
    app = QApplication(sys.argv)
    app.setApplicationName("주의길교회 PPT 도우미 v0.01")
    # Create and show the form
    form = Form()
    form.show()
    # Run the main Qt loop
    sys.exit(app.exec_())

# pyinstaller --name="ppt_helper" -w ppt_helper.py
# pyinstaller --name="ppt_helper" -w ppt_helper.py

# activate powershell
# Set-ExecutionPolicy RemoteSigned
