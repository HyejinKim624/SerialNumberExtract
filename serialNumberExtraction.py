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
        btn3.setText('저장소 선택')
        btn3.clicked.connect(self.clickedButton3)
        btn3.setVisible(False)
        
        global btn4
        btn4 = QPushButton('&Button4', self)
        btn4.setText('일련번호 추출')
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
        
    def clickedButton3(self):   # 파일 저장소 선택
        dialog = QFileDialog(self)
        dialog.setFileMode(QFileDialog.Directory)
        dialog.setViewMode(QFileDialog.Detail)
        dname = None
        if dialog.exec_():
            dname = dialog.selectedFiles()
        global path
        path = dname[0] + '/'
        label.setText(path)
        btn3.setVisible(False)
        btn4.setVisible(True)
        
    def clickedButton4(self):   # 일련번호 추출
        global guideword_df
        i = 1
        
        article_df = pd.read_csv(articlefilename, encoding = 'UTF-8-sig') # 기사 취합본 파일 불러오기
        guideword_df = pd.read_csv(guidewordfilename, encoding = 'UTF-8-sig')    # 색인어 파일 불러오기(일련번호 추출 후 취합은 다시 guideword_df에)
        guideword_list = list(guideword_df) # 색인어를 리스트로
        guideword_df = pd.DataFrame()
        
        for word_list in guideword_list:
            serialnumber = list()  # 색인어를 colname으로 갖는, 추출된 일련번호를 저장할 변수
            word_list = str.split(word_list, '/')   # 슬래시를 기준으로 실제 검색할 단어 리스트 생성
            
            for word in word_list:
                if '(' in word:    # 실제 검색할 단어(word)에서 불필요한 부분 제거 -> (~~~)부분 제거
                    word = word[0 : word.find('(') :]
            
                tmp_list = list()   # word를 포함한 기사의 일련번호를 저장할 리스트
                is_contained = article_df['기사명'].str.contains(word) # 기사명에 word가 포함된 행을 추출하기 위해 인덱스 정보 저장
                tmp_list = article_df[is_contained] # 기사명에 word가 포함된 행만 tmp_list에 저장
            
                if len(tmp_list) == 0:  # word를 포함하는 기사가 없을 경우
                    string = word + '없음'
                    serialnumber.append(string)
                    continue
                serialnumber.extend(list(tmp_list['일련번호'])) # 추출된 데이터프레임에서 '일련번호' 열만 리스트로 만들어 serialnumber에 저장
            
            guideword = str.join('ㆍ', word_list)
            serialnumber = pd.Series(serialnumber, name = guideword)   # 추출된 일련번호 리스트를 시리즈로
            guideword_df = pd.concat([guideword_df, serialnumber], axis = 1)   # 색인어 데이터프레임에 추출된 일련번호를 한 열에 추가
            
            if guideword_df.count().sum() > 5000:  # 데이터 수가 5000을 넘어가면 내보낸 후 이어서 추출
                save_path = path + '색인어' + str(i) + '.csv'
                guideword_df.to_csv(save_path, index = False, encoding = 'UTF-8-sig')
                guideword_df = pd.DataFrame()
                i += 1
        
        article_df = None
        btn4.setEnabled(False)
        label.setText('추출 및 저장 완료')

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MyApp()
    sys.exit(app.exec_())