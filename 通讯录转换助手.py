import re
import sys
import time
import quopri

import pandas
from PyQt5.QtCore import QThread, pyqtSignal, QTime

from PyQt5.QtWidgets import QWidget, QApplication, QGridLayout, QPushButton, QTextBrowser, QFileDialog, QLabel, \
    QVBoxLayout

global getfilenames,targetnames,item
getfilenames = ""
targetnames = ""
item = []

class importevn(QThread):
    triggerd = pyqtSignal()
    def __init__(self):
        super(importevn, self).__init__()
    def run(self):
        global getfilenames
        global item
        if getfilenames:
            with open(getfilenames,"r",encoding="utf-8") as f:
                w = f.readlines()
                j = ""
                for i in w:
                    j = j + i
                    if "END:VCARD" in i:
                        j = j + i
                        item.append(j)
                        j = ""
                f.close()
        self.triggerd.emit()
class SaveXiaomi(QThread):
    trrigerd = pyqtSignal()
    def __init__(self):
        super(SaveXiaomi, self).__init__()
    def run(self):
        name = []
        telnum = []
        for i in item:
            name.append(re.findall("FN:(.*?)\n",i)[0])
            if len(re.findall("TEL;.*?:(.*?)\n",i)) == 1:
                telnum.append(re.findall("TEL;.*?:(.*?)\n",i)[0].replace(" ",""))
            if len(re.findall("TEL;.*?:(.*?)\n",i)) == 0:
                telnum.append("")
            if len(re.findall("TEL;.*?:(.*?)\n",i)) > 1:
                j = ""
                for i in re.findall("TEL;.*?:(.*?)\n",i):
                    j = j + str(i) + ","
                telnum.append(j)
        df = pandas.DataFrame({
            "姓名":name,
            "号码":telnum
        })
        df.to_excel(targetnames,index=None)
        self.trrigerd.emit()
class SaveHuawei(QThread):
    trrigerd = pyqtSignal()
    def __init__(self):
        super(SaveHuawei, self).__init__()
    def run(self):
        global item
        name = []
        telnum = []
        # print(item)
        for i in item:
            if len(re.findall("FN.*?:(.*?)\n",i)) == 1:
                print(re.findall("FN.*?:(.*?)TEL", i.replace("\n", ""))[0].replace("==","="))
                name.append(quopri.decodestring(re.findall("FN.*?:(.*?)TEL", i.replace("\n", ""))[0].replace("==","=").replace(";","").replace("X-ANDROID-CUSTOM","").replace("CHARSET=UTF-8","").replace("vnd.android.cursor.item/nickname","").replace("ENCODING=QUOTED-PRINTABLE","")).decode("utf-8"))
            else:name.append("")
            if len(re.findall("TEL;.*?:(.*?)\n",i)) == 1:
                telnum.append(re.findall("TEL;.*?:(.*?)\n",i)[0].replace(" ",""))
            if len(re.findall("TEL;.*?:(.*?)\n",i)) == 0:
                telnum.append("")
            if len(re.findall("TEL;.*?:(.*?)\n",i)) > 1:
                j = ""
                for i in re.findall("TEL;.*?:(.*?)\n",i):
                    j = j + str(i) + ","
                telnum.append(j)
        df = pandas.DataFrame({
            "姓名":name,
            "号码":telnum
        })
        df.to_excel(targetnames,index=None)
        self.trrigerd.emit()
class Xiaomi(QWidget):
    def __init__(self):
        super(Xiaomi, self).__init__()
        self.initUI()
    def initUI(self):
        self.setWindowTitle("通讯录转换助手")
        self.resize(400,400)
        layout = QGridLayout()
        self.setLayout(layout)
        self.selectbtn = QPushButton("请选择vcf通讯录文件")
        layout.addWidget(self.selectbtn)
        self.savebtn = QPushButton("另存为Excel")
        self.savebtn.setEnabled(False)
        layout.addWidget(self.savebtn)
        self.helpbtn = QPushButton("帮助")
        layout.addWidget(self.helpbtn)
        self.textbrow = QTextBrowser()
        layout.addWidget(self.textbrow)
        self.btnevent()
    def btnevent(self):
        self.selectbtn.clicked.connect(self.filedialog)
        self.savebtn.clicked.connect(self.saveevn)
        self.helpbtn.clicked.connect(lambda :self.textbrow.append("第一步：选择vcf文件\n第二步：点击另存为Excel，并且选择保存路径，提示保存成功即可！\n有任何问题请联系软件作者：QQ:369394074\n"))
    def saveevn(self):
        self.textbrow.append(time.strftime("%Y-%m-%d %H:%M:%S ")+"正在保存，请稍等……\n")
        global targetnames
        targetnames = QFileDialog.getSaveFileName(filter="保存为excel文件(*.xlsx)")[0]
        if targetnames:
            evn = SaveXiaomi()
            ap.append(evn)
            evn.start()
            evn.trrigerd.connect(self.savetedevn)
            self.savebtn.setEnabled(False)
    def savetedevn(self):
        global targetnames
        self.textbrow.append(time.strftime("%Y-%m-%d %H:%M:%S ")+"文件保存成功：%s\n"%(targetnames))
        self.savebtn.setEnabled(True)
    def filedialog(self):
        global getfilenames
        getfilenames = QFileDialog.getOpenFileName(filter="通讯录文件(*.vcf)")[0]
        if getfilenames:
            self.textbrow.append(time.strftime("%Y-%m-%d %H:%M:%S ")+"正在读取通讯录：%s\n"%(getfilenames))
            # timer = QTime()
            # timer.start(100)
            # timer.timeout.connect(lambda :self.textbrow.append(time.strftime("%Y-%m-%d %H:%M:S ")+"正在读取通讯录……"))
            self.selectbtn.setEnabled(False)
            filevn = importevn()
            ap.append(filevn)
            filevn.start()
            filevn.triggerd.connect(self.geteditem)
    def geteditem(self):
        self.selectbtn.setEnabled(True)
        self.savebtn.setEnabled(True)
        self.textbrow.append(time.strftime("%Y-%m-%d %H:%M:%S ",time.localtime())+"通讯录 %s 读取成功\n"%(getfilenames))
class Huawei(QWidget):
    def __init__(self):
        super(Huawei, self).__init__()
        self.initUI()
    def initUI(self):
        self.setWindowTitle("通讯录转换助手")
        self.resize(400,400)
        layout = QGridLayout()
        self.setLayout(layout)
        self.selectbtn = QPushButton("请选择vcf通讯录文件")
        layout.addWidget(self.selectbtn)
        self.savebtn = QPushButton("另存为Excel")
        self.savebtn.setEnabled(False)
        layout.addWidget(self.savebtn)
        self.helpbtn = QPushButton("帮助")
        layout.addWidget(self.helpbtn)
        self.textbrow = QTextBrowser()
        layout.addWidget(self.textbrow)
        self.btnevent()
    def btnevent(self):
        self.selectbtn.clicked.connect(self.filedialog)
        self.savebtn.clicked.connect(self.saveevn)
        self.helpbtn.clicked.connect(lambda :self.textbrow.append("第一步：选择vcf文件\n第二步：点击另存为Excel，并且选择保存路径，提示保存成功即可！\n有任何问题请联系软件作者：QQ:369394074\n"))
    def saveevn(self):
        self.textbrow.append(time.strftime("%Y-%m-%d %H:%M:%S ")+"正在保存，请稍等……\n")
        global targetnames
        targetnames = QFileDialog.getSaveFileName(filter="保存为excel文件(*.xlsx)")[0]
        if targetnames:
            evn = SaveHuawei()
            ap.append(evn)
            evn.start()
            evn.trrigerd.connect(self.savetedevn)
            self.savebtn.setEnabled(False)
    def savetedevn(self):
        global targetnames
        self.textbrow.append(time.strftime("%Y-%m-%d %H:%M:%S ")+"文件保存成功：%s\n"%(targetnames))
        self.savebtn.setEnabled(True)
    def filedialog(self):
        global getfilenames
        getfilenames = QFileDialog.getOpenFileName(filter="通讯录文件(*.vcf)")[0]
        if getfilenames:
            self.textbrow.append(time.strftime("%Y-%m-%d %H:%M:%S ")+"正在读取通讯录：%s\n"%(getfilenames))
            # timer = QTime()
            # timer.start(100)
            # timer.timeout.connect(lambda :self.textbrow.append(time.strftime("%Y-%m-%d %H:%M:S ")+"正在读取通讯录……"))
            self.selectbtn.setEnabled(False)
            filevn = importevn()
            ap.append(filevn)
            filevn.start()
            filevn.triggerd.connect(self.geteditem)
    def geteditem(self):
        self.selectbtn.setEnabled(True)
        self.savebtn.setEnabled(True)
        self.textbrow.append(time.strftime("%Y-%m-%d %H:%M:%S ",time.localtime())+"通讯录 %s 读取成功\n"%(getfilenames))
class Mainwindow(QWidget):
    def __init__(self):
        super(Mainwindow, self).__init__()
        self.initUI()
    def initUI(self):
        self.setWindowTitle("请选择机型")
        self.resize(200,400)
        vlayout = QVBoxLayout()
        self.setLayout(vlayout)
        self.label = QLabel("请选择机型:")
        vlayout.addWidget(self.label)
        self.xiaomibtn = QPushButton("小米")
        vlayout.addWidget(self.xiaomibtn)
        self.huaweibtn = QPushButton("华为")
        vlayout.addWidget(self.huaweibtn)
        vlayout.addStretch()
        self.xiaomibtn.clicked.connect(self.xiaomiwindow)
        self.huaweibtn.clicked.connect(self.huaweiwindow)
    def huaweiwindow(self):
        win = Huawei()
        ap.append(win)
        win.show()
        self.close()
    def xiaomiwindow(self):
        win = Xiaomi()
        ap.append(win)
        win.show()
        self.close()
if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = Mainwindow()
    win.show()
    ap = []
    sys.exit(app.exec_())
