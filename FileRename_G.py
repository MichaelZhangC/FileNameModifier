# -*- coding: utf-8 -*-

import os, sys
import operator
import datetime
import pandas as pd
import shutil
from PyQt5.QtWidgets import *
from PyQt5 import QtCore


class InputDialog(QWidget):
    def __init__(self):
        super().__init__()

        self.initUI()

    def initUI(self):
        self.setGeometry(600, 600, 450, 450)
        self.setWindowTitle('FileRename Toolkit')

        self.btn_Path = QPushButton('Path to Search', self)
        # print(self.btn_Path)
        self.btn_Path.clicked.connect(self.get_search_path)

        self.btn_client = QPushButton('ClientName', self)
        self.btn_client.clicked.connect(self.set_client)

        self.cb_completeness = QCheckBox("completeness", self)
        self.cb_completeness.setChecked(False)

        self.showpath = QLabel(self)
        self.showpath.setText("please enter to path to search")
        self.showpath.setAlignment(QtCore.Qt.AlignCenter)

        self.showclient = QLabel(self)
        self.showclient.setText("xxxx")

        self.period_input = QComboBox(self)
        self.period_input.addItem("q1")
        self.period_input.addItem("q2")
        self.period_input.addItem("q3")
        self.period_input.addItem("final")

        self.namepreview = QLabel(self)
        self.refresh_name()

        self.btn_run = QPushButton('Run', self)
        self.btn_run.clicked.connect(self.runRename)

        self.vbox = QVBoxLayout()
        self.vbox.addWidget(self.btn_Path)
        self.vbox.addWidget(self.showpath)
        self.vbox.addWidget(self.btn_client)
        self.vbox.addWidget(self.showclient)
        self.vbox.addWidget(self.period_input)
        self.vbox.addStretch(1)
        self.vbox.addWidget(self.namepreview)
        self.vbox.addWidget(self.btn_run)

        self.setLayout(self.vbox)

        self.show()

    def get_search_path(self):
        path = QFileDialog.getExistingDirectory(None, "please select folder", "C:/")
        self.showpath.setText(path)

    def refresh_name(self):
        preview_name = "01-" + self.showclient.text() + "-2022-" + self.period_input.currentText().upper() + "-a-" + "1231@0000"
        self.namepreview.setText(preview_name)

    def set_search_path(self):
        text, ok = QInputDialog.getText(self, 'Input Dialog', 'Enter Search Path: ')

        if ok:
            self.showpath.setText(str(text))

    def set_client(self):
        text, ok = QInputDialog.getText(self, 'Input Dialog', 'Enter Client Name: ')

        if ok:
            self.showclient.setText(str(text))
            self.refresh_name()

    def runRename(self):
        if (self.showpath.text() == "please enter to path to search"):
            reply = QMessageBox.warning(self, "Error", "please input necessary information",
                                        QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
        else:
            #print(self.showpath.text(), self.showclient.text(), self.period_input.currentText())
            #print(self.cb_completeness.isChecked())
            rename(self.showpath.text(),
                   self.period_input.currentText(),
                   self.cb_completeness.isChecked(),
                   self.showclient.text())
            reply_get_latest = QMessageBox.information(self, "Option", "Do you want to get the latest 3 file?",
                                        QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
            if reply_get_latest == QMessageBox.Yes:
                get_latest(self.showpath.text())
            message_done = QMessageBox.about(self, "Complete", "Done")


def get_diffdf(df1, df2):
    df = pd.merge(left=df1, right=pd.DataFrame(df2.index, columns=["ID"]), how="left", indicator=True, left_index=True,
                  right_on="ID")
    df = df.loc[df._merge == "left_only", :].drop(columns=["_merge", "ID"])
    return df


def get_duration_name(duration):
    # return the duration in file name format
    year = "2022"
    if duration == "q1":
        return year + "-Q1-"
    if duration == "q2":
        return year + "-interim-"
    if duration == "q3":
        return year + "-Q3-"
    if duration == "final":
        return year + "-final-"
    raise SystemExit("duration format uncorrect!")


def file_identificate(criteria_path):
    criteria = pd.read_excel(criteria_path)


def rename(path, duration, check_request, client=None):
    # 对目录下的文件进行遍历
    durationname = get_duration_name(duration)
    if check_request == False:
        check_request = "n"
    if client == None:
        clientname = "xxxx"
    else:
        clientname = client + "-"
    for file in os.listdir(path):
        filepath = os.path.join(path, file)
        modify_time = datetime.datetime.fromtimestamp(os.path.getmtime(filepath))
        time_atribute = str(modify_time.strftime('%m')) + str(modify_time.strftime('%d')) + "@" + \
                        str(modify_time.strftime('%H')) + str(modify_time.strftime('%M'))
        # 判断文件归属
        if operator.contains(file, "a"):
            newName = "01-" + clientname + durationname + "a-" + time_atribute + ".xlsx"
            os.rename(filepath, os.path.join(path, newName))
        if operator.contains(file, "a") | operator.contains(file.lower(), "p&c") | operator.contains(file.lower(),
                                                                                                      "pc"):
            newName = "02-" + clientname + durationname + "b-" + time_atribute + ".xlsx"
            os.rename(filepath, os.path.join(path, newName))
        elif operator.contains(file, "b") | operator.contains(file.lower(), "life"):
            newName = "03-" + clientname + durationname + "b-" + time_atribute + ".xlsx"
            os.rename(filepath, os.path.join(path, newName))
        elif operator.contains(file, "c") | operator.contains(file.lower(), "health"):
            newName = "04-" + clientname + durationname + "c-" + time_atribute + ".xlsx"
            os.rename(filepath, os.path.join(path, newName))
        elif operator.contains(file, "d") | operator.contains(file.lower(), "amc"):
            newName = "05-" + clientname + durationname + "d-" + time_atribute + ".xlsx"
            os.rename(filepath, os.path.join(path, newName))
        elif operator.contains(file, "e"):
            newName = "06-" + clientname + durationname + "e-" + time_atribute + ".xlsx"
            os.rename(filepath, os.path.join(path, newName))
        elif operator.contains(file, "f"):
            newName = "07-" + clientname + durationname + "f-" + time_atribute + ".xlsx"
            os.rename(filepath, os.path.join(path, newName))
        elif operator.contains(file, "g"):
            newName = "08-" + clientname + durationname + "g-" + time_atribute + ".xlsx"
            os.rename(filepath, os.path.join(path, newName))
        elif operator.contains(file, "h"):
            newName = "10-" + clientname + durationname + "h-" + time_atribute + ".xlsx"
            os.rename(filepath, os.path.join(path, newName))
        elif operator.contains(file, "i"):
            newName = "11-" + clientname + durationname + "i-" + time_atribute + ".xlsx"
            os.rename(filepath, os.path.join(path, newName))
        elif operator.contains(file, "j") | operator.contains(file, "j"):
            newName = "09-" + clientname + durationname + "j-" + time_atribute + ".xlsx"
            os.rename(filepath, os.path.join(path, newName))
        elif operator.contains(file, "k") | operator.contains(file.lower(), "hk"):
            newName = "12-" + clientname + durationname + "k-" + time_atribute + ".xlsx"
            os.rename(filepath, os.path.join(path, newName))


def scanFiles(path_to_search, path_to_save):
    Filedata = {"filePath": [],
                "fileSize": []}
    for file in os.listdir(path_to_search):
        Filedata["filePath"].append(os.path.join(path_to_search, file))
        Filedata["fileSize"].append(os.path.getsize(os.path.join(path_to_search, file)))
    data_df = pd.DataFrame(Filedata)
    # print(data_df)
    path_to_save = path_to_save + '\\' + 'data.csv'
    print(path_to_save)
    data_df.to_csv(path_to_save, encoding="utf_8_sig")


def get_latest(path):
    # this method have to be run after unify the file name
    data = {'filename': [],
            'modifyTime': [],
            'fileClass': [],
            'completeness': []}
    for file in os.listdir(path):
        filepath = os.path.join(path, file)
        # modify_time = datetime.datetime.fromtimestamp(os.path.getmtime(filepath))
        data['filename'].append(file)
        data['modifyTime'].append(file[len(file) - 14:len(file) - 5])
        data['fileClass'].append(file[len(file) - 17:len(file) - 15])
        if "不完整" in file:
            data['completeness'].append(0)
        else:
            data['completeness'].append(1)
    data_df = pd.DataFrame(data)
    data_df_incom = data_df[data_df['completeness'] == 0]
    data_df_com = data_df[data_df['completeness'] == 1]
    data_df_com.sort_values(['fileClass', 'modifyTime'], ascending=[1, 0], inplace=True)
    # print(data_df)
    files_to_keep = data_df_com.groupby(['fileClass']).head(3)
    files_to_keep = files_to_keep.append(data_df_incom)
    # print(files_to_keep)
    files_to_remove = get_diffdf(data_df, files_to_keep)
    # print(files_to_remove)
    des_path = path + r'\legacy'
    if not os.path.exists(des_path):
        os.makedirs(des_path)
    for item in files_to_remove['filename']:
        # print(path + item)
        # print(des_path + item)
        shutil.move(path + '\\' + item, des_path + '\\' + item)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = InputDialog()
    sys.exit(app.exec_())
    # path_input = input("请输入需要更名的文件夹路径")
    # scanFiles(path_input, path_input)
    # savedpath_input = input("请输入需要保存的文件夹路径")
    # period_input = input("请输入本期期间：q1;q2;q3;final")
    # check_request_input = input("请输入是否要确认完整性：y, n")
    # rename(path_input, period_input, check_request_input)
    # filter_request_input = input("请输入是否要自动进行筛选：y, n")
    # if filter_request_input == "y":
    #    get_latest(path_input)
    # 结束
    print("End")
