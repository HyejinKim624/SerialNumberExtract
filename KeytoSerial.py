# 프로그램 작성자 : 김혜진
# 일본학연구소 색인 일련번호 추출 프로그램

import pandas as pd
import copy
import sys
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QDesktopWidget, QFileDialog, QLabel

class MyApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
    
    def initUI(self):
        global label
        global labelText
        labelText = list()
        label = QLabel()
        
        global articleBtn
        articleBtn = QPushButton('&Button1', self)
        articleBtn.setText('기사 파일 선택')
        articleBtn.clicked.connect(self.clickedButton)
              
        global guidewordBtn
        guidewordBtn = QPushButton('&Button2', self)
        guidewordBtn.setText('색인어 파일 선택')
        guidewordBtn.clicked.connect(self.clickedButton2)
        guidewordBtn.setEnabled(False)
        
        global directoryBtn
        directoryBtn = QPushButton('&Button3', self)
        directoryBtn.setText('저장소 선택')
        directoryBtn.clicked.connect(self.clickedButton3)
        directoryBtn.setEnabled(False)
        
        global extractionBtn
        extractionBtn = QPushButton('&Button4', self)
        extractionBtn.setText('일련번호 추출')
        extractionBtn.clicked.connect(self.clickedButton4)
        extractionBtn.setEnabled(False)
        
        vbox = QVBoxLayout()
        vbox.addWidget(articleBtn)
        vbox.addWidget(guidewordBtn)
        vbox.addWidget(directoryBtn)
        vbox.addWidget(extractionBtn)
        vbox.addWidget(label)
        
        self.setLayout(vbox)
        self.setWindowTitle('색인 일련번호 추출 프로그램')
        self.setGeometry(300, 300, 300, 300)
        self.resize(800, 300)
        self.center()
        self.show()

    def center(self):
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())
        
    def clickedButton(self):    # 기사 파일 선택
        global articlefilename
        lt = "기사 파일 : "
        fname = QFileDialog.getOpenFileName(self)

        if fname[0] != "":
            articlefilename = fname[0]
            lt += articlefilename.split('/')[-1]
            labelText.append(lt)
            label.setText("\n".join(labelText))
            articleBtn.setEnabled(False)
            guidewordBtn.setEnabled(True)
    
    def clickedButton2(self):   # 색인어 파일 선택
        global guidewordfilename
        lt = "색인어 파일 : "
        fname = QFileDialog.getOpenFileName(self)

        if fname[0] != "":
            guidewordfilename = fname[0]
            lt += guidewordfilename.split('/')[-1]
            labelText.append(lt)
            label.setText("\n".join(labelText))
            guidewordBtn.setEnabled(False)
            directoryBtn.setEnabled(True)
        
    def clickedButton3(self):   # 파일 저장소 선택
        global path
        lt = "저장소 : "
        dialog = QFileDialog(self)
        dialog.setFileMode(QFileDialog.Directory)
        dialog.setViewMode(QFileDialog.Detail)
        dname = None

        if dialog.exec_():
            dname = dialog.selectedFiles()

        if dname != None and dname[0] != "":
            path = dname[0] + '/'   # 저장소 경로
            lt += path
            labelText.append(lt)
            label.setText("\n".join(labelText))
            directoryBtn.setEnabled(False)
            extractionBtn.setEnabled(True)
        
    def clickedButton4(self):   # 일련번호 추출
        extractionBtn.setEnabled(False)

        i = 1

        article_df = pd.read_excel(articlefilename, dtype = {'일련번호':str}, usecols = "A, G") # 일련번호열과 기사명열 --> 데이터프레임
        guideword_list = list(pd.read_excel(guidewordfilename))                                 # 색인어 --> 리스트

        result = pd.DataFrame()     # 색인어+일련번호가 저장될 데이터프레임
        noResult = pd.DataFrame()   # 그외 색인어가 저장될 데이터프레임

        for guideword in guideword_list:    # 색인어 리스트에서 색인어 1개 꺼내기
            serialnumber_list = list()
            words = str.split(guideword, '/')   # 색인어 분해 --> 단어 리스트

            for word in words:  # 단어 리스트에서 단어 1개 꺼내기
                if '(' in word: # 단어에서 검색에 불필요한 괄호부분 제거
                    word = word[0 : word.find('(')]
                
                is_contained = article_df['기사명']. str.contains(word) # 단어 포함하는 행 정보 저장
                tmp_series = article_df[is_contained]['일련번호']       # 단어 포함하는 일련번호 저장

                if len(tmp_series) == 0:    # 검색 결과가 없으면
                    continue                # 다음 단어로
                
                serialnumber_list.extend(list(tmp_series))  # 색인어에 해당하는 일련번호 저장
                
                del is_contained
                del tmp_series
            
            # 전처리 작업을 거친 색인어를 원래대로 돌리기
            # new_words에서 체크 후 words를 수정하기 위해 복사
            new_words = copy.deepcopy(words)    

            for word in new_words:
                if '(두)' in word:      # (두) 붙은 단어 확인 후
                    words.remove(word)  # words에서 삭제

            new_guideword = str.join('ㆍ', words)   # 각 단어를 중점으로 이어붙여 하나의 색인어로

            del new_words
            del words

            if len(serialnumber_list) == 0: # 검색 결과 없으면 색인어만 1행에 추가
                tmp_series = pd.Series(list(), name = new_guideword, dtype = "str")
                noResult = pd.concat([noResult, tmp_series], axis = 1)
            else: # 걸색 결과가 있는 경우의 처리
                serialnumber_list = list(set(serialnumber_list))    # 중복 제거
                serialnumber_list.sort()    # 오름차순 정렬
                serialnumber_list = pd.Series(serialnumber_list, name = new_guideword, dtype = "float64")
                result = pd.concat([result, serialnumber_list], axis = 1)

            if result.count().sum() > 15000: # 단어 추가 후 result의 데이터 개수 확인해서 너무 크면 내보냄
                save_path = path + '결과 있음' + str(i) + '.xlsx'
                result.to_excel(save_path, index = False, encoding = 'UTF-8')
                del result
                result = pd.DataFrame()
                i += 1

        del article_df

        if result.count().sum() > 0:    # 단어 추가가 모두 완료되고 아직 내보내지 않은 result 내보냄
            save_path = path + '결과 있음' + str(i) + '.xlsx'
            result.to_excel(save_path, index = False, encoding = 'UTF-8')
            del result

        empty_path = path + '결과 없음.xlsx'
        noResult.to_excel(empty_path, index = False, encoding = 'UTF-8')
        
        del noResult
        
        labelText.append("추출 및 저장 완료")
        label.setText("\n".join(labelText))
 
if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MyApp()
    sys.exit(app.exec_())