# 프로그램 작성자 : 김혜진
# 일본학연구소 색인어 일련번호 추출 프로그램

import pandas as pd
import numpy as np
import sys
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QDesktopWidget, QFileDialog, QLabel

class MyApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        
    def initUI(self):
        global label
        label = QLabel()
        
        global btn
        btn = QPushButton('&Button1', self)
        btn.setText('파일 선택')
        btn.clicked.connect(self.clickedButton)
              
        global btn2
        btn2 = QPushButton('&Button2', self)
        btn2.setText('일련번호 추출')
        btn2.clicked.connect(self.clickedButton2)
        btn2.setVisible(False)
        
        vbox = QVBoxLayout()
        vbox.addWidget(btn)
        vbox.addWidget(btn2)
        vbox.addWidget(label)
        
        self.setLayout(vbox)
        self.setWindowTitle('일련번호 추출 프로그램')
        self.setGeometry(300, 300, 300, 300)
        self.resize(400, 150)
        self.center()
        self.show()
        
    def center(self):
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())
        
    def clickedButton(self):
        fname = QFileDialog.getOpenFileName(self)
        label.setText(fname[0])
        btn.setVisible(False)
        btn2.setVisible(True)
    
    def clickedButton2(self):
        articlefilename = label.text()
        global article_df
        article_df = pd.read_csv(articlefilename, encoding = 'UTF-8')
        label.setText('추출 완료')

        
if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MyApp()
    sys.exit(app.exec_())