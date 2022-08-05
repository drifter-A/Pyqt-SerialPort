from serialport_ui import Ui_MainWindow
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QMessageBox
from PyQt5.QtCore import QTimer
from PyQt5.QtCore import pyqtSignal
import serial
import serial.tools.list_ports
import sys
import time
import struct
import pyqtgraph as pg
import openpyxl
from PyQt5.QtWidgets import QFileDialog
from scipy import signal



data_plot = {}
data_plot_list = []
data_plot_list1 = []
data_plot_list2 = []
data_plot_list3 = []

data_plot_list_filt = []
data_plot_list1_filt = []
data_plot_list2_filt = []
data_plot_list3_filt = []

collect_data_start = {}
collect_data = {}
collect_data1 = []
collect_data2 = []


class Pyqt5_serial(QtWidgets.QMainWindow, Ui_MainWindow):
    send_signal = pyqtSignal(int)
    def __init__(self):
        super(Pyqt5_serial, self).__init__()
        self.setupUi(self)
        self.serial_init()
        self.setWindowTitle("串口示波器")
        self.serial_py = serial.Serial()
        pg.setConfigOption("background", "w")

        self.plt_init()
        self.port_check()
        self.change_window_color()
        collect_data_start[0] = 0


    #图表初始化
    def plt_init(self):
        self.mygraph = pg.PlotWidget()
        self.mygraph.setXRange(max = 1000, min = 10)
        self.mygraph.setYRange(max = 4096, min = 0)
        self.mygraph.setLabel(axis='left', text='脉搏')
        self.mygraph.setLabel(axis='top', text='脉搏曲线:红色为曲线1，绿色为曲线2，蓝色为曲线3，紫色为曲线4')
        self.mygraph.setLabel(axis='bottom', text='时间')
        self.curve1 = self.mygraph.plot(pen={'color':(255, 0, 0), 'width':2})
        self.curve2 = self.mygraph.plot(pen={'color':(0, 255, 0), 'width':2})
        self.curve3 = self.mygraph.plot(pen={'color':(0, 0, 255), 'width':2})
        self.curve4 = self.mygraph.plot(pen={'color':(170, 0, 255), 'width':2})
        self.mygraph.enableAutoRange()  # 让整张图始终追踪曲线
        self.formLayout_3.addWidget(self.mygraph)

    #信号和槽连接
    def serial_init(self):
        #串口检测
        self.pushButton_detect_serial.clicked.connect(self.port_check)

        #串口显示
        self.comboBox_serialPort.currentTextChanged.connect(self.port_inf)

        #打开串口按钮
        self.pushButton_connect_serial.clicked.connect(self.port_open_close)

        #发送数据
        self.send_signal.connect(self.data_send)
        self.pushButton_send.clicked.connect(lambda: self.send_signal.emit(0))
        self.pushButton_send_1.clicked.connect(lambda: self.send_signal.emit(1))
        self.pushButton_send_2.clicked.connect(lambda: self.send_signal.emit(2))
        self.pushButton_send_3.clicked.connect(lambda: self.send_signal.emit(3))
        self.pushButton_send_4.clicked.connect(lambda: self.send_signal.emit(4))
        self.pushButton_send_5.clicked.connect(lambda: self.send_signal.emit(5))
        self.pushButton_send_6.clicked.connect(lambda: self.send_signal.emit(6))
        self.pushButton_send_7.clicked.connect(lambda: self.send_signal.emit(7))
        self.pushButton_send_8.clicked.connect(lambda: self.send_signal.emit(8))



        #定时发送数据
        self.timer_send = QTimer()
        self.timer_send.timeout.connect(self.data_send)
        self.lineEdit_t_ms.textChanged.connect(self.data_send_timer)

        #定时接收数据
        self.timer_receive = QTimer(self)
        self.timer_receive.timeout.connect(self.data_receive)

        # 定时绘图
        self.timer_plot = QTimer(self)
        self.timer_plot.timeout.connect(self.dataPlot)

        #清除窗口数据
        self.pushButton_clear.clicked.connect(self.receive_data_clear)

        self.textBrowser_show_text.insertPlainText('欢迎使用智障串口~\r\n')

        #改变窗口的颜色
        self.comboBox_window_color.currentTextChanged.connect(self.change_window_color)

        #保存波形数据
        self.pushButton_save_wave.clicked.connect(self.data_save)

        #清除波形数据
        self.pushButton_clear_wave.clicked.connect(self.data_clear)

        #波形追随
        self.pushButton_wave_follow.clicked.connect(self.wave_follow)

        #开始采集信号
        self.pushButton_start_collect.clicked.connect(self.collect_wave)

        #建立一个线程用于接收信息


    #串口检测
    def port_check(self):
        self.com_dict = {}
        port_list = list(serial.tools.list_ports.comports())
        self.comboBox_serialPort.clear()
        for port in port_list:
            self.com_dict["%s" % port[0]] = "%s" % port[1]
            self.comboBox_serialPort.addItem(port[0])
        if len(self.com_dict) == 0:
            QMessageBox.warning(self, "Port Error", "无串口")

    #串口信息
    def port_inf(self):
        s_inf = self.comboBox_serialPort.currentText()
        if s_inf != "":
            self.pushButton_connect_serial.setEnabled(1)

    #打开或者关闭串口
    def port_open_close(self):
        # 打开串口
        if (self.pushButton_connect_serial.text() == "连接串口"):
            self.serial_py.port = self.comboBox_serialPort.currentText()
            self.serial_py.baudrate = int(self.comboBox_baud.currentText())
            self.serial_py.bytesize = int(self.comboBox_data_bit.currentText())
            self.serial_py.stopbits = int(self.comboBox_stop_bit.currentText())

            if (self.comboBox_check_bit.currentText() == "无校验"):
                self.serial_py.parity = 'N'
            if (self.comboBox_check_bit.currentText() == "奇校验"):
                self.serial_py.parity = 'O'
            if (self.comboBox_check_bit.currentText() == "偶校验"):
                self.serial_py.parity = 'E'

            try:
                self.serial_py.open()
                QMessageBox.warning(self, "连接成功！", "开始干活吧")
            except:
                QMessageBox.critical(self, "连接失败！", "端口选错或者蓝牙未打开")
                return None

            self.pushButton_connect_serial.setText("关闭串口")
            self.comboBox_baud.setDisabled(1)
            self.comboBox_serialPort.setDisabled(1)
            self.comboBox_check_bit.setDisabled(1)
            self.comboBox_stop_bit.setDisabled(1)
            self.comboBox_data_bit.setDisabled(1)
            # 打开定时接收
            self.timer_receive.start(2)
            # 打开定时绘图
            self.timer_plot.start(50)


        #关闭串口
        else:
            self.port_close()

    #关闭串口
    def port_close(self):
        self.timer_send.stop()
        self.timer_receive.stop()
        self.timer_plot.stop()
        try:
            self.serial_py.close()
        except:
            QMessageBox.critical(self, '串口异常', '串口关闭失败，请重启程序！')
            return None

        self.pushButton_connect_serial.setText('连接串口')
        self.comboBox_baud.setEnabled(1)
        self.comboBox_serialPort.setEnabled(1)
        self.comboBox_check_bit.setEnabled(1)
        self.comboBox_stop_bit.setEnabled(1)
        self.comboBox_data_bit.setEnabled(1)

    #发送数据
    def data_send(self, val):
        if self.serial_py.isOpen():
            if val == 0:
                input_s = self.lineEdit.text()
            if val == 1:
                input_s = self.lineEdit_1.text()
            if val == 2:
                input_s = self.lineEdit_2.text()
            if val == 3:
                input_s = self.lineEdit_3.text()
            if val == 4:
                input_s = self.lineEdit_4.text()
            if val == 5:
                input_s = self.lineEdit_5.text()
            if val == 6:
                input_s = self.lineEdit_6.text()
            if val == 7:
                input_s = self.lineEdit_7.text()
            if val == 8:
                input_s = self.lineEdit_8.text()


            if input_s != '':
                # 非空字符串

                if self.checkBox_if_newLine.isChecked():
                    input_s = input_s + '\r\n'
                self.serial_py.write(input_s.encode())
                if self.checkBox_datapause.isChecked() == 0:
                    #时间显示
                    self.textBrowser_show_text.insertPlainText((time.strftime("%H:%M:%S", time.localtime())) + " " + "发送：" + input_s)
                    # 获取text光标
                    textCursor = self.textBrowser_show_text.textCursor()
                    # 滚动到底部
                    textCursor.movePosition(textCursor.End)
                    # 设置光标到text中去
                    self.textBrowser_show_text.setTextCursor(textCursor)


    #接收数据
    def data_receive(self):
            try:
                num = self.serial_py.inWaiting()
            except:
                QMessageBox.warning(self, '串口异常！', '蓝牙可能已关闭')
                self.port_close()
                return None

            else:
                if num > 0:
                    data = self.serial_py.read(num)
                    self.data_analyze(data)

                    if self.checkBox_datapause.isChecked() == 0:
                        self.textBrowser_show_text.insertPlainText((time.strftime("%H:%M:%S", time.localtime())) + " " + "接收：" + data.__str__() + '\r\n')
                        # 获取text光标
                        textCursor = self.textBrowser_show_text.textCursor()
                        # 滚动到底部
                        textCursor.movePosition(textCursor.End)
                        # 设置光标到text中去
                        self.textBrowser_show_text.setTextCursor(textCursor)


    #定时发送数据
    def data_send_timer(self):
        if self.checkBox_circular_send.isChecked():
            self.timer_send.start(int(self.lineEdit_t_ms.text()))
        else:
            self.timer_send.stop()

    #清除显示
    def receive_data_clear(self):
        self.textBrowser_show_text.clear()

    #改变窗口颜色
    def change_window_color(self):
        if self.comboBox_window_color.currentText() == 'whiteblack':
            self.textBrowser_show_text.setStyleSheet("QTextEdit {color:black;background-color:white}")
        elif self.comboBox_window_color.currentText() == 'blackwhite':
            self.textBrowser_show_text.setStyleSheet("QTextEdit {color:white;background-color:black}")
        elif self.comboBox_window_color.currentText() == 'blackgreen':
            self.textBrowser_show_text.setStyleSheet("QTextEdit {color:rgb(0,255,0);background-color:black}")

    #判断是否为波形数据
    def data_analyze(self, data):
        #buf = data.decode('utf-8')
        data_len = len(data)
        for i in range(data_len):
            if (data_len-i) > 21:
                # 数据格式"b,y,16,6,int[4],\r,\n"
                if data[i] == 98:
                    if data[i+1] == 121:
                        if data[i+2]== 16:
                            if data[i+3] == 6:
                                if data[i+20] == 13:
                                    if data[i+21] == 10:
                                        #if len( data[i:i+21].__str__() ) == 81:
                                        x = data[i+4:i+8]
                                        data_plot[0] = struct.unpack('f', x)[0]
                                        data_plot_list.append(data_plot[0])
                                        x = data[i+8:i+12]
                                        data_plot[1] = struct.unpack('f', x)[0]
                                        data_plot_list1.append(data_plot[1])
                                        x = data[i+12:i+16]
                                        data_plot[2] = struct.unpack('f', x)[0]
                                        data_plot_list2.append(data_plot[2])
                                        x = data[i+16:i+20]
                                        data_plot[3] = struct.unpack('f', x)[0]
                                        data_plot_list3.append(data_plot[3])



        if collect_data_start[0] == 1:
            for i in range(data_len):
                if (data_len - i) > 11:
                    #数据格式"c,d,,int[2],\r,\n"
                    if data[i] == 99:
                        if data[i + 1] == 100:
                            if data[i + 10] == 13:
                                if data[i + 11] == 10:
                                    x = data[i + 4:i + 8]
                                    collect_data[0] = struct.unpack('f', x)[0]
                                    collect_data1.append(collect_data[0])
                                    x = data[i + 8:i + 12]
                                    collect_data[1] = struct.unpack('f', x)[0]
                                    collect_data2.append(collect_data[1])
                                    # 'stop'
                                    if data[i + 2] == 99 : # and data[i + 3] == 116 and data[i + 4] == 48 and data[i + 5] == 112 :
                                        book = openpyxl.Workbook()
                                        sheet = book.active
                                        sheet.title = '保存数据'
                                        data_len = collect_data1.__len__()
                                        for i in range(data_len):
                                            sheet.append([collect_data1[i], collect_data2[i]])
                                        filepath, type = QFileDialog.getSaveFileName(self, '数据保存', '/', 'xls(*.xls)')
                                        book.save(filepath)
                                        collect_data_start[0] = 0




    #绘制波形
    def dataPlot(self):

        #滤波，采样频率为3ms
        fs = 3
        b, a = signal.butter(4, [2.0 * 0.8 / fs, 2.0 * 9 / fs], 'bandpass')
        data_plot_list_filt = signal.filtfilt(b, a, data_plot_list)
        data_plot_list1_filt = signal.filtfilt(b, a, data_plot_list1)
        data_plot_list2_filt = signal.filtfilt(b, a, data_plot_list2)
        data_plot_list3_filt = signal.filtfilt(b, a, data_plot_list3)

        if self.checkBox_CH1.isChecked():
            self.curve1.setData(data_plot_list_filt)
        if self.checkBox_CH2.isChecked():
            self.curve2.setData(data_plot_list1_filt)
        if self.checkBox_CH3.isChecked():
            self.curve3.setData(data_plot_list2_filt)
        if self.checkBox_CH4.isChecked():
            self.curve4.setData(data_plot_list3_filt)

        #self.mygraph.plot(data_plot, pen='r')
        #self.mygraph.plot(data_plot_list1, pen='g')
        #self.mygraph.plot(data_plot_list2, pen='b')
        #self.mygraph.plot(data_plot_list3, pen='y')

    #保存数据
    def data_save(self):
        book = openpyxl.Workbook()
        sheet = book.active
        sheet.title = '保存数据'
        data_len = data_plot_list.__len__()
        for i in range(data_len):
            sheet.append([data_plot_list[i], data_plot_list1[i], data_plot_list2[i], data_plot_list3[i] ])
        filepath, type = QFileDialog.getSaveFileName(self, '文件保存', '/', 'xls(*.xls)')
        book.save(filepath)
        #book.save("保存数据.xlsx")

    #清空数据
    def data_clear(self):
        data_plot_list.clear()
        data_plot_list1.clear()
        data_plot_list2.clear()
        data_plot_list3.clear()
        self.curve1.setData(data_plot_list)
        self.curve2.setData(data_plot_list1)
        self.curve3.setData(data_plot_list2)
        self.curve4.setData(data_plot_list3)



    def wave_follow(self):
        self.mygraph.enableAutoRange()  # 让整张图始终追踪曲线

    def collect_wave(self):
        if self.pushButton_start_collect.text() == "开始采集":
            self.pushButton_start_collect.setText("停止采集")
            send_text = 'start_collect 1\r\n'
            collect_data_start[0] = 0
            self.serial_py.write(send_text.encode())
        else:
            self.pushButton_start_collect.setText("开始采集")
            send_text = 'start_collect 0\r\n'
            collect_data_start[0] = 1
            self.serial_py.write(send_text.encode())



if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    myserial = Pyqt5_serial()
    myserial.show()
    sys.exit(app.exec_())
