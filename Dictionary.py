import sys
from openpyxl import load_workbook
from PyQt5.QtWidgets import *
from PyQt5.QtCore import Qt
from PyQt5.QtGui import *
import datetime


class DlgMain(QDialog):
    def __init__(self):
        self.wb = load_workbook('test.xlsx')
        self.sh = self.wb["Лист1"]
        self.start = 65
        self.end = 178
        self.word = ''
        self.fontSize = 14
        self.rand = 0
        self.dateToday = datetime.date.today()
        self.delta = datetime.timedelta(days=2)
        self.dateCell = self.sh['N' + str(self.start)].value.date()
        super().__init__()
        self.setWindowTitle('Dictionary')  # add widgets, set properties
        self.resize(1000, 800)
        self.setStyleSheet("background-color: moccasin;")

        while (self.dateCell - self.dateToday).days > 0:
            self.start += 1
            self.dateCell = self.sh['N' + str(self.start)].value.date()

        if self.start <= self.end:
            pass
        else:
            self.setWindowTitle('ALL WORDS DONE!')

        self.meaningVerb = QLabel('Verb: ' + str(self.sh['E' + str(self.start)].value), self)
        self.meaningVerb.resize(1000, 50)
        self.meaningVerb.move(0, 10)
        self.meaningVerb.setAlignment(Qt.AlignCenter)
        self.meaningVerb.setFont(QFont('Times', self.fontSize))

        self.meaningNoun = QLabel('Noun: ' + str(self.sh['H' + str(self.start)].value), self)
        self.meaningNoun.resize(1000, 50)
        self.meaningNoun.move(0, 60)
        self.meaningNoun.setAlignment(Qt.AlignCenter)
        self.meaningNoun.setFont(QFont('Times', self.fontSize))

        self.meaningOther = QLabel('Other: ' + str(self.sh['K' + str(self.start)].value), self)
        self.meaningOther.resize(1000, 50)
        self.meaningOther.move(0, 120)
        self.meaningOther.setAlignment(Qt.AlignCenter)
        self.meaningOther.setFont(QFont('Times', self.fontSize))

        self.ledText = QLineEdit('', self)  # Enter text
        self.ledText.resize(400, 20)
        self.ledText.move(300, 220)
        self.ledText.setAlignment(Qt.AlignCenter)

        self.checkBtn = QPushButton('Check', self)
        self.checkBtn.resize(100, 20)
        self.checkBtn.move(350, 270)
        self.checkBtn.clicked.connect(self.check_word)

        self.answerBtn = QPushButton('???', self)
        self.answerBtn.resize(80, 20)
        self.answerBtn.move(460, 270)
        self.answerBtn.clicked.connect(self.answer)

        self.continueBtn = QPushButton('Next', self)
        self.continueBtn.resize(100, 20)
        self.continueBtn.move(550, 270)
        self.continueBtn.clicked.connect(self.next_word)

        self.answer = QLabel('Result', self)
        self.answer.resize(200, 30)
        self.answer.move(400, 310)
        self.answer.setAlignment(Qt.AlignCenter)
        self.answer.setFont(QFont('Times', 14))

        self.exampleVerb = QLabel('Verb: ', self)
        self.exampleVerb.resize(1000, 30)
        self.exampleVerb.move(0, 350)
        self.exampleVerb.setAlignment(Qt.AlignCenter)
        self.exampleVerb.setFont(QFont('Times', self.fontSize))

        self.exampleNoun = QLabel('Noun: ', self)
        self.exampleNoun.resize(1000, 30)
        self.exampleNoun.move(0, 450)
        self.exampleNoun.setAlignment(Qt.AlignCenter)
        self.exampleNoun.setFont(QFont('Times', self.fontSize))

        self.exampleOther = QLabel('Other: ', self)
        self.exampleOther.resize(1000, 30)
        self.exampleOther.move(0, 550)
        self.exampleOther.setAlignment(Qt.AlignCenter)
        self.exampleOther.setFont(QFont('Times', self.fontSize))

        self.translate = QLabel('', self)
        self.translate.resize(400, 30)
        self.translate.move(300, 650)
        self.translate.setAlignment(Qt.AlignCenter)
        self.translate.setFont(QFont('Times', 14))

    def check_word(self):

        self.word = str(self.sh['C' + str(self.start)].value)
        if self.ledText.text() == self.word:
            self.answer.setText('Good!')
            self.exampleVerb.setText('Verb: ' + str(self.sh['F' + str(self.start)].value))
            self.exampleNoun.setText('Noun: ' + str(self.sh['I' + str(self.start)].value))
            self.exampleOther.setText('Other: ' + str(self.sh['L' + str(self.start)].value))
            self.sh['N' + str(self.start)] = self.dateToday + self.delta
            self.translate.setText(str(self.sh['P' + str(self.start)].value))
            self.wb.save(filename='test.xlsx')
        else:
            self.answer.setText('Wrong!')
            self.translate.setText('')

    def next_word(self):
        self.start += 1
        self.dateCell = self.sh['N' + str(self.start)].value.date()
        while (self.dateCell - self.dateToday).days > 0:
            self.start += 1
            self.dateCell = self.sh['N' + str(self.start)].value.date()

        if self.start <= self.end:
            pass
        else:
            self.setWindowTitle('ALL WORDS DONE!')

        self.answer.setText('Result')
        self.meaningVerb.setText('Verb: ' + str(self.sh['E' + str(self.start)].value))
        self.meaningNoun.setText('Noun: ' + str(self.sh['H' + str(self.start)].value))
        self.meaningOther.setText('Other: ' + str(self.sh['K' + str(self.start)].value))
        self.exampleVerb.setText('Verb: ')
        self.exampleNoun.setText('Noun: ')
        self.exampleOther.setText('Other: ')
        self.translate.setText('')
        self.ledText.setText('')

    def answer(self):
        self.answer.setText(str(self.sh['C' + str(self.start)].value))
        self.exampleVerb.setText('Verb: ' + str(self.sh['F' + str(self.start)].value))
        self.exampleNoun.setText('Noun: ' + str(self.sh['I' + str(self.start)].value))
        self.exampleOther.setText('Other: ' + str(self.sh['L' + str(self.start)].value))
        self.translate.setText(str(self.sh['P' + str(self.start)].value))
        

if __name__ == '__main__':
    app = QApplication(sys.argv)
    dlgMain = DlgMain()
    dlgMain.show()
    sys.exit(app.exec_())
