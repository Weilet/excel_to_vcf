import os
from PyQt5.uic import loadUi
from PyQt5.QtWidgets import QMainWindow, QApplication, QFileDialog
from excf import VcfCreator


class Window(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('excf')
        self.resize(800, 800)
        self.move(400, 200)
        self.set_ui()
        self.showButton.hide()
        self.worker = VcfCreator()

    def set_ui(self):
        loadUi('ui/excf.ui', self)

    def openExcel(self):
        fileName, fileType = QFileDialog.getOpenFileName(self, "选取文件", "./", "excel(*.xls, *.xlsx)")
        try:
            self.worker.read_excel(fileName)
            self.worker.create_vcf()
            self.worker.filepath = os.path.abspath('contact.vcf')
            self.showButton.show()
        except (IOError, NameError, TypeError) as e:
            print(e)

    def showVcard(self):
        os.system('explorer {}'.format(self.worker.filepath))


