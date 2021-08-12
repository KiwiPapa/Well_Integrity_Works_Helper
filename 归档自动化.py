import sys
from win32 import win32api, win32gui, win32print
from win32.lib import win32con
from win32.win32api import GetSystemMetrics
import win32api
import win32con
import pywintypes
from auto_archive import Ui_MainWindow
from PyQt5.Qt import QPoint, QPropertyAnimation, QEasingCurve, QAbstractAnimation
from PyQt5 import QtCore, QtGui, QtWidgets, QtNetwork
from PyQt5.QtCore import QDate, QTime, QBasicTimer, QDateTime, Qt, QTimer
from PyQt5.QtPrintSupport import QPageSetupDialog, QPrintDialog, QPrinter
from PyQt5.QtWidgets import (QApplication, QColorDialog, QDialog, QFileDialog,
                             QFontDialog, QLabel, QLineEdit, QMainWindow,
                             QMessageBox, QPushButton, QRadioButton,
                             QTableWidgetItem, QTextEdit, QWidget, QStyleFactory)

class Main_window(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super(Main_window, self).__init__()
        self.setupUi(self)
        self.statusBar().showMessage('Ready')
        self.setWindowOpacity(0.98)
        self.setObjectName("mainWindow")
        # 归档自动化
        ###################################################
        self.pushButton.clicked.connect(self.change_screen_dpi)
        self.pushButton_2.clicked.connect(self.auto_launch_archive_system)
        self.pushButton_3.clicked.connect(self.xiaoduishigongxinxi_jibenxinxi_autofill)

    def get_real_resolution(self):
        """获取真实的分辨率"""
        hDC = win32gui.GetDC(0)
        # 横向分辨率
        w = win32print.GetDeviceCaps(hDC, win32con.DESKTOPHORZRES)
        # 纵向分辨率
        h = win32print.GetDeviceCaps(hDC, win32con.DESKTOPVERTRES)
        return w, h


    def get_screen_size(self):
        """获取缩放后的分辨率"""
        w = GetSystemMetrics(0)
        h = GetSystemMetrics(1)
        return w, h


    def change_screen_dpi(self):
        # 改变分辨率
        real_resolution = self.get_real_resolution()
        screen_size = self.get_screen_size()
        print(real_resolution)
        print(screen_size)

        screen_scale_rate = round(real_resolution[0] / screen_size[0], 2)
        print(screen_scale_rate)

        devmode = pywintypes.DEVMODEType()

        devmode.PelsWidth = 1980
        devmode.PelsHeight = 1080
        # devmode.PelsWidth = 2560
        # devmode.PelsHeight = 1440

        devmode.Fields = win32con.DM_PELSWIDTH | win32con.DM_PELSHEIGHT

        win32api.ChangeDisplaySettings(devmode, 0)


    def auto_launch_archive_system(self):
        # 调取鼠标工具
        # pyautogui.mouseInfo()

        # while True:
        #     p = pyautogui.position()
        #     print(p)
        #     time.sleep(0.2)

        # b = pyautogui.locateOnScreen('111.png')
        # print(b)

        wh = pyautogui.size()
        screen_width = int(wh.width)
        screen_height = int(wh.height)

        app_dir = r'D:\Program Files (x86)\Well Logging DAU App Beta\logging_app.exe'
        try:
            os.startfile(app_dir)
            time.sleep(2)

            fw = pyautogui.getActiveWindow()
            print(fw.title)

            fw = pyautogui.getAllTitles()
            for item in fw:
                if '井场一体化' in item:
                    print(item)
                    break
            fw1 = pyautogui.getWindowsWithTitle(item)[0]
            print('The title of windows is: ' + fw1.title)
            fw1.maximize()

            # pyautogui.moveTo(1362/2560*screen_width, 699/1440*screen_height, duration=0.5)
            pyautogui.click(int(1362 / 2560 * screen_width), int(699 / 1440 * screen_height))
            pyautogui.write("yangy_cj")
            pyautogui.click(int(1362 / 2560 * screen_width), int(733 / 1440 * screen_height))
            pyautogui.write("123")
            pyautogui.click(int(1339 / 2560 * screen_width), int(789 / 1440 * screen_height))
        except:
            pass


    def xiaoduishigongxinxi_jibenxinxi_autofill(self):
        wh = pyautogui.size()
        screen_width = int(wh.width)
        screen_height = int(wh.height)

        driller_depth = self.lineEdit_4.text()
        pyautogui.click(int(309 / 2560 * screen_width), int(416 / 1440 * screen_height))
        pyautogui.write(driller_depth)

        log_depth = self.lineEdit_53.text()
        pyautogui.click(int(896 / 2560 * screen_width), int(414 / 1440 * screen_height))
        pyautogui.write(log_depth)

        pou_mian = "碳酸盐岩"
        pyperclip.copy(pou_mian)
        pyautogui.click(int(2038 / 2560 * screen_width), int(416 / 1440 * screen_height))
        # pyautogui.write(pou_mian)
        # pyperclip.paste()
        pyautogui.hotkey('ctrl', 'v')

        max_deviation = self.lineEdit_24.text()
        pyautogui.click(int(309 / 2560 * screen_width), int(458 / 1440 * screen_height))
        pyautogui.write(max_deviation)

        max_deviation_depth = self.lineEdit_28.text()
        pyautogui.click(int(896 / 2560 * screen_width), int(456 / 1440 * screen_height))
        pyautogui.write(max_deviation_depth)
        pass

if __name__ == "__main__":
    # 运行主程序
    app = QApplication(sys.argv)
    QApplication.setStyle(QStyleFactory.create("Fusion"))
    main = Main_window()
    main.show()
    sys.exit(app.exec_())