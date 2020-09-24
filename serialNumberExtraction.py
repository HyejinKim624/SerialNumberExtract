# 프로그램 작성자 : 김혜진
# 일본학연구소 색인어 일련번호 추출 프로그램

import pandas as pd
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
        btn.setText('기사 파일 선택')
        btn.clicked.connect(self.clickedButton)
              
        global btn2
        btn2 = QPushButton('&Button2', self)
        btn2.setText('색인어 파일 선택')
        btn2.clicked.connect(self.clickedButton2)
        btn2.setVisible(False)
        
        global btn3
        btn3 = QPushButton('&Button3', self)
        btn3.setText('일련번호 추출')
        btn3.clicked.connect(self.clickedButton3)
        btn3.setVisible(False)
        
        global btn4
        btn4 = QPushButton('&Button3', self)
        btn4.setText('파일 저장소 선택')
        btn4.clicked.connect(self.clickedButton4)
        btn4.setVisible(False)
        
        vbox = QVBoxLayout()
        vbox.addWidget(btn)
        vbox.addWidget(btn2)
        vbox.addWidget(btn3)
        vbox.addWidget(btn4)
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
        
    def clickedButton(self):    # 기사 파일 선택
        fname = QFileDialog.getOpenFileName(self)
        label.setText(fname[0])
        global articlefilename
        articlefilename = label.text()
        btn.setVisible(False)
        btn2.setVisible(True)
    
    def clickedButton2(self):   # 색인어 파일 선택
        fname = QFileDialog.getOpenFileName(self)
        label.setText(fname[0])
        global guidewordfilename
        guidewordfilename = label.text()
        btn2.setVisible(False)
        btn3.setVisible(True)
        
    def clickedButton3(self):   # 일련번호 추출
        label.setText('추출 중')
        
        global guideword_df
        
        article_df = pd.read_csv(articlefilename, encoding = 'UTF-8') # 기사 취합본 파일 불러오기
        guideword_df = pd.read_csv(guidewordfilename, encoding = 'UTF-8')    # 색인어 파일 불러오기(일련번호 추출 후 취합은 다시 guideword_df에)
        guideword_list = list(guideword_df) # 색인어를 리스트로
        guideword_df = pd.DataFrame()
        
        for word_list in guideword_list:
            serialnumber = list()  # 색인어를 colname으로 갖는, 추출된 일련번호를 저장할 변수
            guideword = str(word_list)
            if '(' in word_list:  # 색인어에서 불필요한 부분 제거 -> (~~)부분 제거
                word_list = word_list[0 : word_list.find('(') :]
            word_list = str.split(word_list, 'ㆍ')   # 불필요한 괄호 부분 제거 후 중점(ㆍ)을 기준으로 실제 검색할 단어를 추출
            for word in word_list:
                tmp_list = list()  # 색인어를 바탕으로 추출한 일련번호를 저장할 리스트
                is_contained = article_df['기사명'].str.contains(word) # 실제 검색할 단어를 포함한 기사를 추려내기 위한 is_contained
                tmp_list = article_df[is_contained] # 실제 검색할 단어(word_list의 원소들)를 포함한 기사(데이터 프레임)를 추출
                if len(tmp_list) == 0:  # word로 검색한 결과가 0개일 경우
                    string = word + '없음'
                    serialnumber.append(string)
                    continue
                serialnumber.extend(list(tmp_list['일련번호'])) # word_list에서 각각 추출된 일련번호 취합  
            serialnumber = pd.Series(serialnumber, name = guideword)    # 일련번호 리스트를 시리즈로
            guideword_df = pd.concat([guideword_df, serialnumber], axis = 1)    # 색인어 데이터프레임에 일련번호 추가
        
        article_df = None
        btn3.setVisible(False)
        btn4.setVisible(True)
        label.setText('추출 완료')

    def clickedButton4(self):   # 파일 저장소 선택
        dialog = QFileDialog(self)
        dialog.setFileMode(QFileDialog.Directory)
        dialog.setViewMode(QFileDialog.Detail)
        if dialog.exec_():
            global dname
            dname = dialog.selectedFiles()
        output = dname[0] + '/일련번호추출파일.csv'
        guideword_df.to_csv(output, header = True, index = False, encoding = 'UTF-8-sig') #csv파일로 내보내기
        btn4.setEnabled(False)
        label.setText('저장 완료')
        
if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MyApp()
    sys.exit(app.exec_())
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
       
   


