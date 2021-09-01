import pandas as pd
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QMessageBox
import os

if_data_source = False
data_source = '检测人员花名册.xlsx'

if_header_line = False
if if_header_line is False:
    header_line = 0
else:
    pass

count = 1
total_num = 0


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        self.dialog = MainWindow
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(1241, 854)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(390, 40, 541, 71))
        self.label.setStyleSheet("font: 32pt \"微软雅黑\";\n"
"color: rgb(170, 85, 255);")
        self.label.setObjectName("label")
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(730, 740, 501, 71))
        self.label_2.setStyleSheet("font: 16pt \"微软雅黑\";\n"
"color: rgb(0, 170, 255);")
        self.label_2.setObjectName("label_2")
        self.default_path_button = QtWidgets.QPushButton(self.centralwidget)
        self.default_path_button.setGeometry(QtCore.QRect(60, 160, 231, 91))
        self.default_path_button.setStyleSheet("font: 20pt \"华文琥珀\";")
        self.default_path_button.setObjectName("default_path_button")
        self.set_path_button = QtWidgets.QPushButton(self.centralwidget)
        self.set_path_button.setGeometry(QtCore.QRect(340, 160, 231, 91))
        self.set_path_button.setStyleSheet("font: 20pt \"华文琥珀\";")
        self.set_path_button.setObjectName("set_path_button")
        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        self.label_3.setGeometry(QtCore.QRect(80, 240, 1141, 621))
        self.label_3.setText("")
        self.label_3.setPixmap(QtGui.QPixmap("pic/pic2.jpg"))
        self.label_3.setObjectName("label_3")
        self.output_unfinished_button = QtWidgets.QPushButton(self.centralwidget)
        self.output_unfinished_button.setGeometry(QtCore.QRect(930, 160, 251, 91))
        self.output_unfinished_button.setStyleSheet("font: 20pt \"华文琥珀\";")
        self.output_unfinished_button.setObjectName("output_unfinished_button")
        self.output_finished_button = QtWidgets.QPushButton(self.centralwidget)
        self.output_finished_button.setGeometry(QtCore.QRect(620, 160, 251, 91))
        self.output_finished_button.setStyleSheet("font: 20pt \"华文琥珀\";")
        self.output_finished_button.setObjectName("output_finished_button")
        self.label_3.raise_()
        self.label.raise_()
        self.label_2.raise_()
        self.default_path_button.raise_()
        self.set_path_button.raise_()
        self.output_unfinished_button.raise_()
        self.output_finished_button.raise_()
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 1241, 26))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        self.output_finished_button.clicked.connect(self.generate_finished_name_list)
        self.output_unfinished_button.clicked.connect(self.generate_unfinished_name_list)
        self.set_path_button.clicked.connect(self.set_path)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def generate_finished_name_list(self):
        file_list = os.listdir('./信息表')
        index_list = []
        for file in file_list:
            try:
                char = file.split('.')[0][0]
            except Exception as e:
                continue
            if char == '第':
                df = pd.read_excel(os.path.join(r'./信息表', file))
                for index in df['序号']:
                    index_list.append(index)
        data = pd.read_excel(data_source, header=header_line)
        df_list = []
        for index in index_list:
            df_list.append(data[data['序号'].isin([index])])
        output_df = pd.concat([i for i in df_list], axis=0)
        output_df.to_excel('信息表/完成检测居民基本信息表.xlsx'.format(count), encoding='utf-8', index=False)
        QMessageBox.information(self.dialog, "Information", '成功生成已完成人名单!')

    def generate_unfinished_name_list(self):
        file_list = os.listdir('./信息表')
        index_list = []
        for file in file_list:
            try:
                char = file.split('.')[0][0]
            except Exception as e:
                continue
            if char == '第':
                df = pd.read_excel(os.path.join(r'./信息表', file))
                for index in df['序号']:
                    index_list.append(index)
        data = pd.read_excel(data_source, header=header_line)
        df_list = []
        for i in range(0, len(data) + 1):
            if i not in index_list:
                df_list.append(data[data['序号'].isin([i])])
        output_df = pd.concat([i for i in df_list], axis=0)
        output_df.to_excel('信息表/未完成检测居民基本信息表.xlsx'.format(count), encoding='utf-8', index=False)
        QMessageBox.information(self.dialog, "Information", '成功生成未完成人名单!')

    def set_path(self):
        global if_data_source
        if_data_source = True

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.label.setText(_translate("MainWindow", "社区防疫工作辅助系统"))
        self.label_2.setText(_translate("MainWindow", "制作者：益民园居委会志愿者 张雪健"))
        self.default_path_button.setText(_translate("MainWindow", "默认路径模式"))
        self.set_path_button.setText(_translate("MainWindow", "设置路径模式"))
        self.output_unfinished_button.setText(_translate("MainWindow", "输出未完成名单"))
        self.output_finished_button.setText(_translate("MainWindow", "输出已完成名单"))


class Setpath_Ui_Dialog(object):
    def setupUi(self, Dialog):
        self.path_text = ''
        self.index_text = ''
        self.dialog = Dialog
        Dialog.setObjectName("Dialog")
        Dialog.resize(1258, 864)
        self.label = QtWidgets.QLabel(Dialog)
        self.label.setGeometry(QtCore.QRect(50, -30, 541, 211))
        self.label.setStyleSheet("font: 36pt \"华文琥珀\";\n"
"color: rgb(170, 85, 255);")
        self.label.setObjectName("label")
        self.label_2 = QtWidgets.QLabel(Dialog)
        self.label_2.setGeometry(QtCore.QRect(120, 460, 941, 61))
        self.label_2.setStyleSheet("font: 20pt \"Agency FB\";")
        self.label_2.setObjectName("label_2")
        self.lineEdit = QtWidgets.QLineEdit(Dialog)
        self.lineEdit.setGeometry(QtCore.QRect(120, 550, 891, 51))
        self.lineEdit.setObjectName("lineEdit")
        self.clear_button = QtWidgets.QPushButton(Dialog)
        self.clear_button.setGeometry(QtCore.QRect(340, 650, 101, 31))
        self.clear_button.setObjectName("clear_button")
        self.confirm_button = QtWidgets.QPushButton(Dialog)
        self.confirm_button.setGeometry(QtCore.QRect(120, 650, 101, 31))
        self.confirm_button.setObjectName("confirm_button")
        self.return_button = QtWidgets.QPushButton(Dialog)
        self.return_button.setGeometry(QtCore.QRect(1070, 790, 93, 28))
        self.return_button.setObjectName("return_button")
        self.label_3 = QtWidgets.QLabel(Dialog)
        self.label_3.setGeometry(QtCore.QRect(120, 170, 871, 61))
        self.label_3.setStyleSheet("font: 20pt \"Agency FB\";")
        self.label_3.setObjectName("label_3")
        self.lineEdit_2 = QtWidgets.QLineEdit(Dialog)
        self.lineEdit_2.setGeometry(QtCore.QRect(120, 260, 891, 51))
        self.lineEdit_2.setObjectName("lineEdit_2")
        self.clear_button_2 = QtWidgets.QPushButton(Dialog)
        self.clear_button_2.setGeometry(QtCore.QRect(340, 360, 101, 31))
        self.clear_button_2.setObjectName("clear_button_2")
        self.confirm_button_2 = QtWidgets.QPushButton(Dialog)
        self.confirm_button_2.setGeometry(QtCore.QRect(120, 360, 101, 31))
        self.confirm_button_2.setObjectName("confirm_button_2")
        self.change_count_button = QtWidgets.QPushButton(Dialog)
        self.change_count_button.setGeometry(QtCore.QRect(830, 790, 93, 28))
        self.change_count_button.setObjectName("change_count_button")

        self.retranslateUi(Dialog)
        self.confirm_button_2.clicked.connect(self.get_path_text)
        self.confirm_button.clicked.connect(self.generate_excel)
        self.clear_button.clicked.connect(self.lineEdit.clear)
        self.clear_button_2.clicked.connect(self.lineEdit_2.clear)
        self.return_button.clicked.connect(Dialog.close)
        QtCore.QMetaObject.connectSlotsByName(Dialog)

    def get_path_text(self):
        self.path_text = self.lineEdit_2.text()

    def get_path(self):
        global data_source
        if if_data_source is False:
            data_source = '街道办事处全员核酸检测居民基本信息登记表.xlsx'
        else:
            try:
                data_source = self.text
            except e as Exception:
                QMessageBox.warning(self.dialog, "Warning", "请先输入文件路径并点击确认！")

    def generate_excel(self):
        global count, total_num
        data = pd.read_excel(data_source, header=header_line)
        self.index_text = self.lineEdit.text()
        index_list = self.index_text.split(' ')
        df_list = []
        for index in index_list:
            index = int(index)
            df_list.append(data[data['序号'].isin([index])])
            total_num += 1
        len_index_list = len(index_list)
        output_df = pd.concat([i for i in df_list], axis=0)
        output_df.to_excel('信息表/第{}组检测居民基本信息表.xlsx'.format(count), encoding='utf-8', index=False)
        QMessageBox.information(self.dialog, "Information", "已成功生成第{}组信息表，该组人数为{}，目前已完成排队登记{}人！".format(count, len_index_list, total_num))
        count += 1

    def retranslateUi(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "Dialog"))
        self.label.setText(_translate("Dialog", "设置路径模式"))
        self.label_2.setText(_translate("Dialog", "请在下面的空白框中输入该组检测居民的序号（用一个空格隔开）"))
        self.clear_button.setText(_translate("Dialog", "重新输入"))
        self.confirm_button.setText(_translate("Dialog", "确认"))
        self.return_button.setText(_translate("Dialog", "返回上一级"))
        self.label_3.setText(_translate("Dialog", "请在下面的空白框中输入要读入的文件路径（包括扩展名）"))
        self.clear_button_2.setText(_translate("Dialog", "重新输入"))
        self.confirm_button_2.setText(_translate("Dialog", "确认"))
        self.change_count_button.setText(_translate("Dialog", "更改计数器"))


class Defaultpath_Ui_Dialog(object):
    def setupUi(self, Dialog):
        self.index_text = ''
        self.dialog = Dialog
        Dialog.setObjectName("Dialog")
        Dialog.resize(1258, 864)
        self.label = QtWidgets.QLabel(Dialog)
        self.label.setGeometry(QtCore.QRect(50, -30, 541, 211))
        self.label.setStyleSheet("font: 36pt \"华文琥珀\";\n"
"color: rgb(170, 85, 255);")
        self.label.setObjectName("label")
        self.label_2 = QtWidgets.QLabel(Dialog)
        self.label_2.setGeometry(QtCore.QRect(120, 170, 991, 61))
        self.label_2.setStyleSheet("font: 20pt \"Agency FB\";")
        self.label_2.setObjectName("label_2")
        self.lineEdit = QtWidgets.QLineEdit(Dialog)
        self.lineEdit.setGeometry(QtCore.QRect(120, 260, 891, 51))
        self.lineEdit.setObjectName("lineEdit")
        self.clear_button = QtWidgets.QPushButton(Dialog)
        self.clear_button.setGeometry(QtCore.QRect(340, 360, 101, 31))
        self.clear_button.setObjectName("clear_button")
        self.confirm_button = QtWidgets.QPushButton(Dialog)
        self.confirm_button.setGeometry(QtCore.QRect(120, 360, 101, 31))
        self.confirm_button.setObjectName("confirm_button")
        self.return_button = QtWidgets.QPushButton(Dialog)
        self.return_button.setGeometry(QtCore.QRect(1070, 790, 93, 28))
        self.return_button.setObjectName("return_button")
        self.change_count_button = QtWidgets.QPushButton(Dialog)
        self.change_count_button.setGeometry(QtCore.QRect(830, 790, 93, 28))
        self.change_count_button.setObjectName("change_count_button")

        self.retranslateUi(Dialog)
        self.confirm_button.clicked.connect(self.generate_excel)
        self.clear_button.clicked.connect(self.lineEdit.clear)
        self.return_button.clicked.connect(Dialog.close)
        QtCore.QMetaObject.connectSlotsByName(Dialog)

    def generate_excel(self):
        global count, total_num
        data = pd.read_excel(data_source, header=header_line)
        self.index_text = self.lineEdit.text()
        index_list = self.index_text.split(' ')
        total_num_list = []
        df_list = []
        for index in index_list:
            index_int = int(index)
            df_list.append(data[data['序号'].isin([index_int])])
            total_num += 1
            total_num_list.append(total_num)
        len_index_list = len(index_list)
        output_df = pd.concat([i for i in df_list], axis=0)
        output_df['核酸检测序号'] = total_num_list
        output_df.to_excel(excel_writer='信息表/第{}组检测居民基本信息表.xlsx'.format(count), encoding='utf-8', index=False)
        QMessageBox.information(self.dialog, "Information", "已成功生成第{}组信息表，该组人数为{}，目前已完成排队登记{}人！".format(count, len_index_list, total_num))
        count += 1

    def retranslateUi(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "Dialog"))
        self.label.setText(_translate("Dialog", "默认路径模式"))
        self.label_2.setText(_translate("Dialog", "请在下面的空白框中输入该组检测居民的序号（用一个空格隔开）"))
        self.clear_button.setText(_translate("Dialog", "重新输入"))
        self.confirm_button.setText(_translate("Dialog", "确认"))
        self.return_button.setText(_translate("Dialog", "返回上一级"))
        self.change_count_button.setText(_translate("Dialog", "更改计数器"))


class Count_Ui_Dialog(object):
    def setupUi(self, Dialog):
        self.text = ''
        self.dialog = Dialog
        Dialog.setObjectName("Dialog")
        Dialog.resize(457, 353)
        self.label = QtWidgets.QLabel(Dialog)
        self.label.setGeometry(QtCore.QRect(30, 20, 351, 71))
        self.label.setStyleSheet("font: 16pt \"华文琥珀\";")
        self.label.setObjectName("label")
        self.lineEdit = QtWidgets.QLineEdit(Dialog)
        self.lineEdit.setGeometry(QtCore.QRect(40, 130, 141, 31))
        self.lineEdit.setObjectName("lineEdit")
        self.pushButton = QtWidgets.QPushButton(Dialog)
        self.pushButton.setGeometry(QtCore.QRect(240, 130, 93, 28))
        self.pushButton.setObjectName("pushButton")
        self.pushButton_2 = QtWidgets.QPushButton(Dialog)
        self.pushButton_2.setGeometry(QtCore.QRect(320, 300, 93, 28))
        self.pushButton_2.setObjectName("pushButton_2")

        self.retranslateUi(Dialog)
        self.pushButton.clicked.connect(self.change_count)
        self.pushButton_2.clicked.connect(Dialog.close)
        QtCore.QMetaObject.connectSlotsByName(Dialog)

    def change_count(self):
        global count, total_num
        self.text = self.lineEdit.text()
        count = int(self.text)
        file_list = os.listdir('./信息表')
        index_list = []
        for file in file_list:
            try:
                char = file.split('.')[0][0]
            except Exception as e:
                continue
            if char == '第':
                df = pd.read_excel(os.path.join(r'./信息表', file))
                for index in df['序号']:
                    index_list.append(index)
        total_num = len(index_list)

    def retranslateUi(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "Dialog"))
        self.label.setText(_translate("Dialog", "请输入更改后的计数器的值"))
        self.pushButton.setText(_translate("Dialog", "确认"))
        self.pushButton_2.setText(_translate("Dialog", "返回"))

if __name__ == '__main__':
    import sys

    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    main_ui = Ui_MainWindow()
    main_ui.setupUi(MainWindow)

    # 默认路径窗口
    Default_Dialog = QtWidgets.QDialog()
    default_ui = Defaultpath_Ui_Dialog()
    default_ui.setupUi(Default_Dialog)

    # 设置路径窗口
    Setpath_Dialog = QtWidgets.QDialog()
    setpath_ui = Setpath_Ui_Dialog()
    setpath_ui.setupUi(Setpath_Dialog)

    # 更改计数器窗口
    Count_Dialog = QtWidgets.QDialog()
    count_ui = Count_Ui_Dialog()
    count_ui.setupUi(Count_Dialog)

    # 设置窗口链接调用
    setpath_ui.change_count_button.clicked.connect(Count_Dialog.show)
    default_ui.change_count_button.clicked.connect(Count_Dialog.show)
    main_ui.default_path_button.clicked.connect(Default_Dialog.show)
    main_ui.set_path_button.clicked.connect(Setpath_Dialog.show)

    MainWindow.show()
    sys.exit(app.exec_())
