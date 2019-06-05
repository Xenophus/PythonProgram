import sys
import xlwt
from xlutils.copy import copy
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from PyQt5.QtWidgets import *
from GAAP import *

class MainWindow(QWidget):

    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setGeometry(200, 200, 1000, 800)
        self.setWindowTitle('DSF跑数工具')
        self.setWindowIcon(QIcon('gaap.png'))

        '''主程序区'''
        group1 = QGroupBox('主程序区', self)
        group1.setGeometry(20, 20, 960, 180)

        label0 = QLabel('选择计算区间：', self)
        label0.move(50, 50)
        label1 = QLabel('开始时间：', self)
        label1.move(50, 100)

        # 开始年份
        self.startYear = QComboBox(self)
        self.startYear.addItems(['2015年', '2016年', '2017年', '2018年', '2019年', '2020年',
                                 '2021年', '2022年', '2023年', '2024年', '2025年', '2026年'])
        # startYear.move(150, 50)
        self.startYear.setCurrentIndex(3)
        self.startYear.setGeometry(QRect(125, 95, 75, 23))

        # 开始月份
        self.startMonth = QComboBox(self)
        self.startMonth.addItems(['01月', '02月', '03月', '04月', '05月', '06月',
                                  '07月', '08月', '09月', '10月', '11月', '12月'])
        self.startMonth.setCurrentIndex(0)
        self.startMonth.setGeometry(QRect(225, 95, 75, 23))

        label2 = QLabel('结束时间：', self)
        label2.move(50, 150)

        # 结束年份
        self.endYear = QComboBox(self)
        self.endYear.addItems(['2015年', '2016年', '2017年', '2018年', '2019年', '2020年',
                               '2021年', '2022年', '2023年', '2024年', '2025年', '2026年'])
        self.endYear.setCurrentIndex(3)
        self.endYear.setGeometry(QRect(125, 145, 75, 23))

        # 结束月份
        self.endMonth = QComboBox(self)
        self.endMonth.addItems(['01月', '02月', '03月', '04月', '05月', '06月',
                                '07月', '08月', '09月', '10月', '11月', '12月'])
        self.endMonth.setCurrentIndex(11)
        self.endMonth.setGeometry((QRect(225, 145, 75, 23)))

        # 生成还款计划表复选框
        self.schdule = QCheckBox('生成还款计划表', self)
        self.schdule.setCheckState(Qt.Unchecked)
        self.schdule.move(500, 50)
        # 生成收入报表复选框
        self.report = QCheckBox('生成收入报表', self)
        self.report.setCheckState(Qt.Checked)
        self.report.move(500, 100)

        # 选择进程数
        label12 = QLabel('选择进程数：', self)
        label12.move(700, 53)
        self.process = QComboBox(self)
        self.process.addItems(['1', '2', '4', '6', '8'])
        self.process.setCurrentIndex(2)
        self.process.setGeometry((QRect(800, 49, 75, 23)))

        # 开始按钮
        run = QPushButton('开始跑数', self)
        run.move(500, 150)
        run.clicked.connect(self.onClick)

        quitmain = QPushButton('退出程序', self)
        quitmain.clicked.connect(QCoreApplication.instance().quit)
        quitmain.resize(quitmain.sizeHint())
        quitmain.move(600, 150)

        '''业务参数区'''
        group2 = QGroupBox('业务参数区', self)
        group2.setGeometry(20, 220, 960, 550)

        label3 = QLabel('UPFRONT_COST：', self)
        label3.move(50, 250)
        label4 = QLabel('ONGOING_COST：', self)
        label4.move(50, 320)
        label5 = QLabel('MATCH_EARLY_RATE：', self)
        label5.move(50, 390)
        label6 = QLabel('TPY_Model_Ratio1：', self)
        label6.move(50, 460)
        label7 = QLabel('TPY_Model_Ratio2：', self)
        label7.move(50, 530)
        label10 = QLabel('TPY_Model_Ratio3：', self)
        label10.move(50, 600)
        label11 = QLabel('TPY_Model_Ratio4：', self)
        label11.move(50, 670)

        self.upfront = QLineEdit(self)
        self.upfront.move(200, 245)
        self.upfront.setEnabled(False)

        self.ongoing = QLineEdit(self)
        self.ongoing.move(200, 315)
        self.ongoing.setEnabled(False)

        self.matchEarlyRate = QLineEdit(self)
        self.matchEarlyRate.move(200, 385)
        self.matchEarlyRate.setEnabled(False)

        self.TPYModelRatio1 = QLineEdit(self)
        self.TPYModelRatio1.move(200, 455)
        self.TPYModelRatio1.setEnabled(False)

        self.TPYModelRatio2 = QLineEdit(self)
        self.TPYModelRatio2.move(200, 525)
        self.TPYModelRatio2.setEnabled(False)

        self.TPYModelRatio3 = QLineEdit(self)
        self.TPYModelRatio3.move(200, 595)
        self.TPYModelRatio3.setEnabled(False)

        self.TPYModelRatio4 = QLineEdit(self)
        self.TPYModelRatio4.move(200, 665)
        self.TPYModelRatio4.setEnabled(False)

        label8 = QLabel('Implied Price Concession Ratio 对照表：', self)
        label8.move(450, 250)
        #implied_price_concession_ratio 对照表
        self.ipcr = QTableWidget(self)
        self.ipcr.setColumnCount(2)
        self.ipcr.setRowCount(38)
        self.ipcr.horizontalHeader().setVisible(False)
        self.ipcr.verticalHeader().setVisible(False)
        self.ipcr.setGeometry(450, 275, 220, 210)
        self.ipcr.setEditTriggers(QTableWidget.NoEditTriggers)

        label9 = QLabel('LossRatio and Margin：', self)
        label9.move(450, 500)
        #loss_ratio_and_margin 参数表
        self.lrm = QTableWidget(self)
        self.lrm.setColumnCount(5)

        self.lrm.setGeometry(445, 525, 520, 220)
        self.lrm.horizontalHeader().setVisible(False)
        self.lrm.verticalHeader().setVisible(False)
        self.lrm.setEditTriggers(QTableWidget.NoEditTriggers)

        #显示参数
        self.showWindow()

        #修改参数按钮
        change = QPushButton('修改参数', self)
        change.move(875, 350)
        change.clicked.connect(self.changeParamter)
        # 保存按钮
        save = QPushButton('保存修改', self)
        save.move(875, 400)
        save.clicked.connect(self.saveChange)

        self.show()

    def getPeriod(self):
        startDate = self.startYear.currentText()[:-1] + '-' +self.startMonth.currentText()[:-1]
        endDate = self.endYear.currentText()[:-1] +  '-' + self.endMonth.currentText()[:-1]
        period = startDate + '-' + endDate
        return [period]

    #检查复选框，需要导出哪种报表
    def runChecked(self):
        if self.schdule.checkState() == 0:
            config['OUTPUT_SCHEDULE'] = 0
        else:
            config['OUTPUT_SCHEDULE'] = 1
        if self.report.checkState() == 0:
            config['OUTPUT_REPORT'] = 0
        else:
            config['OUTPUT_REPORT'] = 1
        process = self.process.currentText()
        config['PROCESS'] = int(process)

    @pyqtSlot()
    # 开始运行事件
    def onClick(self):
        period = self.getPeriod()
        PERIOD = set_period(period)
        self.runChecked()
        self.th = WorkThread(PERIOD=PERIOD)
        self.th.start()

    '''获取业务参数的更改日志'''
    @staticmethod
    def getChangeLog():
        path = os.getcwd() + os.sep + 'config.xls'
        excel = xlrd.open_workbook(path)
        busConfig = excel.sheet_by_index(0)
        changeLog = excel.sheet_by_index(4)
        count = changeLog.nrows

        dicList = [{}, {}, {}, {}, {}, {}, {}]

        for i in range(1, count):
            time = xlrd.xldate_as_datetime(changeLog.cell_value(i, 0), 0).strftime('%Y/%m/%d')
            time = datetime.datetime.strptime(time, '%Y/%m/%d')

            for j in range(1, 8):
                try:
                    if changeLog.cell_value(i, j) != '':
                        dicList[j - 1][time] = changeLog.cell_value(i, j)
                except:
                    pass

        ever = '2015/01/01'
        ever = datetime.datetime.strptime(ever, '%Y/%m/%d')
        t = 0
        for r in range(0, 7):
            dicList[t][ever] = busConfig.cell_value(r, 1)
            t += 1

        return dicList


        # 获得配置文件中的参数

    #读取implied_price_concession_ratio和loss_ratio and margin
    @staticmethod
    def configParamter():
        ipcrList = []
        lrmList = []
        path = os.getcwd() + os.sep + 'config.xls'
        try:
            excel = xlrd.open_workbook(path)
            sheet2 = excel.sheet_by_index(2)
            rows2 = sheet2.nrows
            cols2 = sheet2.ncols
            for i in range(0, rows2):
                ipcrList.append(sheet2.cell_value(i, cols2-1))
            ipcrList = [round(i*100, 8) for i in ipcrList]
            ipcrList = [str(i) + '%' for i in ipcrList]

            sheet3 = excel.sheet_by_index(3)
            rows3 = sheet3.nrows
            for i in range(1, rows3):
                lrm = []
                for j in range(0, 5):
                    if j in [3, 4]:
                        lrm.append(str(round(sheet3.cell_value(i, j) * 100, 6)) + '%')
                    elif j in [0, 1]:
                        lrm.append(xlrd.xldate_as_datetime(sheet3.cell_value(i, j), 0).strftime("%Y/%m/%d"))
                    else:
                        lrm.append(sheet3.cell_value(i, j))
                lrmList.append(lrm)

        except:
            print('config.xls file error.')
            exit()
        return ipcrList, lrmList

    def showWindow(self):
        paramList = [self.upfront, self.ongoing, self.matchEarlyRate, self.TPYModelRatio1,
                     self.TPYModelRatio2, self.TPYModelRatio3, self.TPYModelRatio4]
        changeLog = self.getChangeLog()

        for i in range(0, len(changeLog)):
            dic = changeLog[i]
            key = sorted(dic.keys(), reverse=True)[0]
            paramList[i].setText(str(dic[key]))

        ipcrList, lrmList = self.configParamter()
        ipcrRows = len(ipcrList) + 1
        self.ipcr.setRowCount(ipcrRows)
        self.ipcr.setItem(0, 0, QTableWidgetItem('Month of Lean'))
        self.ipcr.setItem(0, 1, QTableWidgetItem('IPCR'))
        for i in range(0, 37):
            self.ipcr.setItem(i+1, 0, QTableWidgetItem(str(i)))
            self.ipcr.setItem(i+1, 1, QTableWidgetItem(ipcrList[i]))

        lrmRows = len(lrmList) + 2
        self.lrm.setRowCount(lrmRows)
        titleList = ['start_date', 'end_date', 'sornum', 'annual_loss_ratio', 'annual_margin']
        for i in range(0, 5):
            self.lrm.setItem(0, i, QTableWidgetItem(titleList[i]))
        for i in range(1, lrmRows-1):
            for j in range(0, 5):
                self.lrm.setItem(i, j, QTableWidgetItem(lrmList[i-1][j]))

    @pyqtSlot()
    #修改参数点击事件
    def changeParamter(self):
        self.upfront.setEnabled(True)
        self.ongoing.setEnabled(True)
        self.matchEarlyRate.setEnabled(True)
        self.TPYModelRatio1.setEnabled(True)
        self.TPYModelRatio2.setEnabled(True)
        self.TPYModelRatio3.setEnabled(True)
        self.TPYModelRatio4.setEnabled(True)
        self.ipcr.setEditTriggers(QTableWidget.AllEditTriggers)
        self.lrm.setEditTriggers(QTableWidget.AllEditTriggers)
        for i in range(0, 2):
            self.ipcr.item(0, i).setFlags(Qt.ItemIsEnabled)
        for i in range(1, 37):
            self.ipcr.item(i, 0).setFlags(Qt.ItemIsEnabled)
        for i in range(0, 5):
            self.lrm.item(0, i).setFlags(Qt.ItemIsEnabled)

    @pyqtSlot()
    #保存参数点击事件
    def saveChange(self):
        self.saveParamter()
        self.saveTable1()
        self.saveTable2()
        self.showWindow()
        self.upfront.setEnabled(False)
        self.ongoing.setEnabled(False)
        self.matchEarlyRate.setEnabled(False)
        self.TPYModelRatio1.setEnabled(False)
        self.TPYModelRatio2.setEnabled(False)
        self.TPYModelRatio3.setEnabled(False)
        self.TPYModelRatio4.setEnabled(False)
        self.ipcr.setEditTriggers(QTableWidget.NoEditTriggers)
        self.lrm.setEditTriggers(QTableWidget.NoEditTriggers)

        '''
            刷新表格参数
        '''

    #业务参数修改
    def saveParamter(self):
        dicList = self.getChangeLog()
        paramList = []
        for i in range(0, len(dicList)):
            dic = dicList[i]
            key = sorted(dic.keys(), reverse=True)[0]
            paramList.append(dic[key])

        newParam = []
        today = datetime.datetime.today().strftime('%Y/%m/%d')
        today = datetime.datetime.strptime(today, '%Y/%m/%d')
        newParam.append(today)

        '''
            编辑框参数修改
        '''
        if float(self.upfront.text()) == paramList[0]:
            newParam.append('')
        else:
            newParam.append(float(self.upfront.text()))
        if float(self.ongoing.text()) == paramList[1]:
            newParam.append('')
        else:
            newParam.append(float(self.ongoing.text()))
        if float(self.matchEarlyRate.text()) == paramList[2]:
            newParam.append('')
        else:
            newParam.append(float(self.matchEarlyRate.text()))
        if float(self.TPYModelRatio1.text()) == paramList[3]:
            newParam.append('')
        else:
            newParam.append(float(self.TPYModelRatio1.text()))
        if float(self.TPYModelRatio2.text()) == paramList[4]:
            newParam.append('')
        else:
            newParam.append(float(self.TPYModelRatio2.text()))
        if float(self.TPYModelRatio3.text()) == paramList[5]:
            newParam.append('')
        else:
            newParam.append(float(self.TPYModelRatio3.text()))
        if float(self.TPYModelRatio4.text()) == paramList[6]:
            newParam.append('')
        else:
            newParam.append(float(self.TPYModelRatio4.text()))

        # 判断业务参数是否发生改动
        sign1 = False
        for i in newParam:
            if i != '' and i != today:
                sign1 = True
                break

        # #如果参数列表不全为空，就写入更改日志
        style = xlwt.XFStyle()
        style.num_format_str = 'yyyy/mm/dd'
        path = os.getcwd() + os.sep + 'config.xls'
        if sign1:
            oldExcel = xlrd.open_workbook(path, formatting_info=True)
            newExcel = copy(oldExcel)
            changeLog = oldExcel.sheet_by_index(4)
            rows4 = changeLog.nrows
            wt4 = newExcel.get_sheet(4)
            if rows4 > 1:
                oldDate = changeLog.cell_value(rows4 - 1, 0)
                oldDate = xlrd.xldate_as_datetime(oldDate, 0).strftime('%Y/%m/%d')
                oldDate = datetime.datetime.strptime(oldDate, '%Y/%m/%d')
                if oldDate == today:
                    for i in range(0, len(newParam)):
                        if i == 0:
                            wt4.write(rows4 - 1, i, newParam[i], style)
                        else:
                            wt4.write(rows4 - 1, i, newParam[i])
                else:
                    for i in range(0, len(newParam)):
                        if i == 0:
                            wt4.write(rows4, i, newParam[i], style)
                        else:
                            wt4.write(rows4, i, newParam[i])

            newExcel.save('config.xls')

    #表格参数修改
    def saveTable1(self):
        ipcrList, lrmList = self.configParamter()
        sign2 = False
        signList = []
        for i in range(1, 38):
            if self.ipcr.item(i, 1).text() != ipcrList[i - 1]:
                sign2 = True
                signList.append(i)

        style = xlwt.XFStyle()
        style.num_format_str = 'yyyy/mm/dd'
        path = os.getcwd() + os.sep + 'config.xls'
        if sign2:
            oldExcel = xlrd.open_workbook(path, formatting_info=True)
            newExcel = copy(oldExcel)
            try:
                wt2 = newExcel.get_sheet(2)
                for i in signList:
                    temp = self.ipcr.item(i, 1).text()
                    temp = float(temp[: -1]) / 100
                    wt2.write(i - 1, 1, temp)
            except:
                print('error')
            newExcel.save('config.xls')

    def saveTable2(self):
        lrmCount = self.lrm.rowCount()
        sign3 = True
        try:
            for i in range(0, 5):
                if self.lrm.item(lrmCount - 1, i).text() == '':
                    sign3 = False
        except:
            sign3 = False
        if sign3:
            try:
                path = os.getcwd() + os.sep + 'config.xls'
                oldExcel = xlrd.open_workbook(path, formatting_info=True)
                newExcel = copy(oldExcel)
                style = xlwt.XFStyle()
                style.num_format_str = 'yyyy/mm/dd'
                wt3 = newExcel.get_sheet(3)
                for i in range(0, 5):
                    if i in [0, 1]:
                        time = datetime.datetime.strptime(self.lrm.item(lrmCount-1, i).text(), '%Y/%m/%d')
                        wt3.write(lrmCount-1, i, time, style)
                    if i == 2:
                        wt3.write(lrmCount-1, i, self.lrm.item(lrmCount-1, i).text())
                    if i in [3, 4]:
                        wt3.write(lrmCount-1, i, float(self.lrm.item(lrmCount-1, i).text()[:-1]) / 100)
                newExcel.save('config.xls')
            except:
                pass

'''工作线程'''
class WorkThread(QThread):
    finishSingal = pyqtSignal()

    def __init__(self,PERIOD, parent=None):
        super(WorkThread, self).__init__(parent)
        self.PERIOD = PERIOD

    def run(self):
        PERIOD = self.PERIOD
        if config['PROCESS'] == 1:
            #单进程处理
            single_task(PERIOD, 1)
        else:
            #多进程处理
            multi_task(PERIOD, config['PROCESS'])
        self.finishSingal.emit()

'''进度条'''
# class ProgressBar(QWidget):
#     def __init__(self):
#         QWidget.__init__(self)
#
#         self.setGeometry(300, 300, 250, 150)
#         self.setWindowTitle('ProgressBar')
#         self.pbar = QProgressBar(self)
#         self.pbar.setGeometry(30, 40, 200, 25)
#
#         self.button = QPushButton('Start', self)
#         self.button.setFocusPolicy(Qt.NoFocus)
#         self.button.move(40, 80)
#
#         self.button.clicked.connect(self.onStart)
#         self.timer = QBasicTimer()
#         self.step = 0
#
#     def timerEvent(self, event):
#         if self.step >= 100:
#             self.timer.stop()
#             return
#         self.step = self.step + 1
#         self.pbar.setValue(self.step)
#
#     def onStart(self):
#         if self.timer.isActive():
#             self.timer.stop()
#             self.button.setText('Start')
#         else:
#             self.timer.start(100, self)
#             self.button.setText('Stop')

if __name__ == '__main__':
    app = QApplication(sys.argv)
    mw = MainWindow()
    sys.exit(app.exec_())