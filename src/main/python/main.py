import scrapy
from scrapy.crawler import CrawlerProcess
from twisted.internet import reactor
import xlsxwriter
import xml.etree.ElementTree as ET
import threading
import time
import xlrd
import socket
from twisted.internet.task import deferLater
from scrapy.utils.project import get_project_settings
from pyqtspinner.spinner import WaitingSpinner
from fbs_runtime.application_context.PyQt5 import ApplicationContext
from PyQt5.QtWidgets import (QApplication,QGridLayout,QDialog,QGroupBox, QHBoxLayout, QLabel, QLineEdit,
                             QStyleFactory, QTableWidget, QTextEdit,QPushButton,QRadioButton,QCheckBox,
                             QVBoxLayout, QWidget,QDesktopWidget,QTableWidgetItem,QCompleter,QMessageBox)
from PyQt5.QtGui import QColor
from PyQt5.QtCore import Qt
import sys
import winsound



class CodalSpider(scrapy.Spider):
    name = "codal"

    def start_requests(self):
        urls = [
            'https://search.codal.ir/api/search/v2/q?&Audited=true&AuditorRef=-1&Category=-1&Childs=true&CompanyState=-1&CompanyType=-1&Consolidatable=true&IsNotAudited=false&Length=-1&LetterType=-1&Mains=true&NotAudited=true&NotConsolidatable=true&PageNumber=1&Publisher=false&TracingNo=-1&search=false',
        ]
        for url in urls:
            yield scrapy.Request(url=url, callback=self.parse)

    def parse(self, response):
        root = ET.fromstring(response.body)
        letter=root[0]
        CodalLetterHeaderDtoElements=[]
        Widget_GUI_obj=Widget_GUI.getInstance()
        for CodalLetterHeaderDto in letter:
            CodalLetterHeaderDtoElements.append(CodalLetterHeaderDto)
        for i in range(0,20):
            Widget_GUI_obj.addDataToTableWidget(CodalLetterHeaderDtoElements[19-i][11].text,CodalLetterHeaderDtoElements[19-i][1].text,CodalLetterHeaderDtoElements[19-i][12].text,CodalLetterHeaderDtoElements[19-i][8].text)



class Widget_GUI(QDialog):
    __instance = None
    def __init__(self, parent=None):
        super(Widget_GUI, self).__init__(parent)


        self.process = CrawlerProcess(get_project_settings())


        self.counter=0
        self.stopFlag=False
        self.secondRunFlag=False
        self.exitFlag=False
        self.addedData=[]
        self.showFullScreen()
        self.setFixedSize(1000, 600)#214, 245, 245
        self.setStyleSheet("background-color: rgb(214, 245, 245);")

        self.input_symbols_text=[]
        self.symbols_textInput_objs=[]
        self.input_keywords_text=[]
        self.keywords_textInput_objs=[]


        #************************************************
        #************************************************
        # SPINNER
        self.spinner = WaitingSpinner(self)
        self.spinner.setNumberOfLines(100)
        self.spinner.setMinimumTrailOpacity(0.0)
        self.spinner.setLineLength(5.0)
        self.spinner.setLineWidth(10.0)
        self.spinner.setInnerRadius(40.0)
        self.spinner.setRevolutionsPerSecond(1.4)
        self.spinner.setColor(QColor(0, 0, 0, 255))
        #************************************************
        #************************************************

        self.tableWidget = QTableWidget(1,4)
        self.tableWidget.horizontalHeader().hide()
        self.tableWidget.setColumnWidth(3,150)
        self.tableWidget.setColumnWidth(2,225)
        self.tableWidget.setColumnWidth(1,360)
        self.tableWidget.setColumnWidth(0,200)
        item=self.tableWidget.setItem(0,3, QTableWidgetItem("نماد"))
        self.tableWidget.setItem(0,2, QTableWidgetItem("نام شرکت"))
        self.tableWidget.setItem(0,1, QTableWidgetItem("عنوان اطلاعیه"))
        self.tableWidget.setItem(0,0, QTableWidgetItem("زمان انتشار"))


        
        item = self.tableWidget.item(0, 3)
        item.setTextAlignment(Qt.AlignCenter)
        item = self.tableWidget.item(0, 2)
        item.setTextAlignment(Qt.AlignCenter)
        item = self.tableWidget.item(0, 1)
        item.setTextAlignment(Qt.AlignCenter)
        item = self.tableWidget.item(0, 0)
        item.setTextAlignment(Qt.AlignCenter)
        self.tableWidget.setStyleSheet("background-color: white;")

        self.createTopLeftGroupBox()
        self.createTopMiddleGroupBox()
        self.createTopRightGroupBox()

        self.all_radioButton = QRadioButton("all")
        self.all_radioButton.toggled.connect(self.topLeftGroupBox.setDisabled)
        self.selective_radioButton = QRadioButton("selective")
        self.selective_radioButton.setChecked(True)

        self.with_keyword_checkbox = QCheckBox("keyword")
        self.with_keyword_checkbox.setChecked(True)
        self.with_keyword_checkbox.stateChanged.connect(self.topMiddleGroupBox.setEnabled)

        topLayout = QHBoxLayout()
        topLayout.addWidget(self.all_radioButton)
        topLayout.addWidget(self.selective_radioButton)
        topLayout.addWidget(self.with_keyword_checkbox)
        topLayout.addStretch(1)

        mainLayout = QGridLayout()
        mainLayout.addLayout(topLayout, 0, 0, 1, 3)
        mainLayout.addWidget(self.topLeftGroupBox, 1, 0)
        mainLayout.addWidget(self.topMiddleGroupBox, 1, 1)
        mainLayout.addWidget(self.topRightGroupBox, 1, 2)
        mainLayout.addWidget(self.tableWidget, 2, 0,3,3)
        mainLayout.setRowStretch(1, 1)
        mainLayout.setColumnStretch(0, 1)
        mainLayout.setColumnStretch(1, 1)
        mainLayout.setColumnStretch(2, 1)
        self.setLayout(mainLayout)

        self.setWindowTitle("Codal_Crawler")
        self.center()
        self.changeStyle('windowsvista')

        if Widget_GUI.__instance != None:
            raise Exception("This class is a singleton!")
        else:
            Widget_GUI.__instance = self


    def getInstance():
        if Widget_GUI.__instance == None:
            Widget_GUI()
        return Widget_GUI.__instance

    def center(self):
        qr = self.frameGeometry()
        # center point of screen
        cp = QDesktopWidget().availableGeometry().center()
        # move rectangle's center point to screen's center point
        qr.moveCenter(cp)
        # top left of rectangle becomes top left of window centering it
        self.move(qr.topLeft())

    def changeStyle(self, styleName):
        QApplication.setStyle(QStyleFactory.create(styleName))
        self.changePalette()

    def changePalette(self):
        QApplication.setPalette(QApplication.style().standardPalette())

    def createTopLeftGroupBox(self):
        self.topLeftGroupBox = QGroupBox("")
        self.topLeftGroupBox.setObjectName("ColoredGroupBox")
        self.topLeftGroupBox.setStyleSheet("QGroupBox#ColoredGroupBox { border: 1px solid black;}")

        symbolsLabel = QLabel("Symbols")
        Symbol1 = QLineEdit('')
        Symbol2 = QLineEdit('')
        Symbol3 = QLineEdit('')
        Symbol4 = QLineEdit('')
        Symbol5 = QLineEdit('')
        Symbol6 = QLineEdit('')
        Symbol7 = QLineEdit('')
        Symbol8 = QLineEdit('')
        Symbol9 = QLineEdit('')
        Symbol10 = QLineEdit('')

        Symbol1.setStyleSheet("background-color: white;")
        Symbol2.setStyleSheet("background-color: white;")
        Symbol3.setStyleSheet("background-color: white;")
        Symbol4.setStyleSheet("background-color: white;")
        Symbol5.setStyleSheet("background-color: white;")
        Symbol6.setStyleSheet("background-color: white;")
        Symbol7.setStyleSheet("background-color: white;")
        Symbol8.setStyleSheet("background-color: white;")
        Symbol9.setStyleSheet("background-color: white;")
        Symbol10.setStyleSheet("background-color: white;")

        self.symbols_textInput_objs.append(Symbol1)
        self.symbols_textInput_objs.append(Symbol2)
        self.symbols_textInput_objs.append(Symbol3)
        self.symbols_textInput_objs.append(Symbol4)
        self.symbols_textInput_objs.append(Symbol5)
        self.symbols_textInput_objs.append(Symbol6)
        self.symbols_textInput_objs.append(Symbol7)
        self.symbols_textInput_objs.append(Symbol8)
        self.symbols_textInput_objs.append(Symbol9)
        self.symbols_textInput_objs.append(Symbol10)


        #***************************************************************
        # Auto Complete

        autocomplete_symbols=[]

        loc = ("symbols.xlsx") 
        wb = xlrd.open_workbook(loc) 
        sheet = wb.sheet_by_index(0)
  
        for i in range(sheet.nrows): 
            autocomplete_symbols.append(sheet.cell_value(i, 0))

        completer = QCompleter(autocomplete_symbols)
        Symbol1.setCompleter(completer)
        Symbol2.setCompleter(completer)
        Symbol3.setCompleter(completer)
        Symbol4.setCompleter(completer)
        Symbol5.setCompleter(completer)
        Symbol6.setCompleter(completer)
        Symbol7.setCompleter(completer)
        Symbol8.setCompleter(completer)
        Symbol9.setCompleter(completer)
        Symbol10.setCompleter(completer)

        #***************************************************************
        

        layout = QVBoxLayout()
        layout.addWidget(symbolsLabel)
        layout.addWidget(Symbol1)
        layout.addWidget(Symbol2)
        layout.addWidget(Symbol3)
        layout.addWidget(Symbol4)
        layout.addWidget(Symbol5)
        layout.addWidget(Symbol6)
        layout.addWidget(Symbol7)
        layout.addWidget(Symbol8)
        layout.addWidget(Symbol9)
        layout.addWidget(Symbol10)
        layout.addStretch(1)
        self.topLeftGroupBox.setLayout(layout)


    def createTopMiddleGroupBox(self):
        self.topMiddleGroupBox = QGroupBox("")
        self.topMiddleGroupBox.setObjectName("ColoredGroupBox")
        self.topMiddleGroupBox.setStyleSheet("QGroupBox#ColoredGroupBox { border: 1px solid black;}")

        keywordsLabel = QLabel("Keywords")
        keyword1 = QLineEdit('')
        keyword2 = QLineEdit('')
        keyword3 = QLineEdit('')
        keyword4 = QLineEdit('')
        keyword5 = QLineEdit('')

        keyword1.setStyleSheet("background-color: white;")
        keyword2.setStyleSheet("background-color: white;")
        keyword3.setStyleSheet("background-color: white;")
        keyword4.setStyleSheet("background-color: white;")
        keyword5.setStyleSheet("background-color: white;")

        self.keywords_textInput_objs.append(keyword1)
        self.keywords_textInput_objs.append(keyword2)
        self.keywords_textInput_objs.append(keyword3)
        self.keywords_textInput_objs.append(keyword4)
        self.keywords_textInput_objs.append(keyword5)

        layout = QVBoxLayout()
        layout.addWidget(keywordsLabel)
        layout.addWidget(keyword1)
        layout.addWidget(keyword2)
        layout.addWidget(keyword3)
        layout.addWidget(keyword4)
        layout.addWidget(keyword5)
        layout.addStretch(1)
        self.topMiddleGroupBox.setLayout(layout)

    def createTopRightGroupBox(self):
        self.topRightGroupBox = QGroupBox("")
        self.topRightGroupBox.setObjectName("ColoredGroupBox")
        self.topRightGroupBox.setStyleSheet("QGroupBox#ColoredGroupBox { border: 1px solid black;}")

        goPushButton = QPushButton("Go")
        goPushButton.setDefault(False)
        goPushButton.setStyleSheet("background-color: rgb(153, 255, 187);")#Green
        goPushButton.clicked.connect(self.buttonGoClicked)

        clearPushButton = QPushButton("Clear OutPut")
        clearPushButton.setDefault(False)
        clearPushButton.setStyleSheet("background-color: lightGray;")
        clearPushButton.clicked.connect(self.buttonClearClicked)

        stopPushButton = QPushButton("Stop Crawling")
        stopPushButton.setDefault(False)
        stopPushButton.setStyleSheet("background-color: rgb(255, 179, 255);")
        stopPushButton.clicked.connect(self.buttonStopClicked)

        exitPushButton = QPushButton("Exit")
        exitPushButton.setDefault(False)
        exitPushButton.setStyleSheet("background-color: rgb(255, 179, 255);")
        exitPushButton.clicked.connect(self.buttonExitClicked)


        layout = QVBoxLayout()
        layout.addWidget(goPushButton)
        layout.addWidget(clearPushButton)
        layout.addWidget(stopPushButton)
        layout.addWidget(exitPushButton)
        layout.addWidget(self.spinner)
        layout.addStretch(1)
        self.topRightGroupBox.setLayout(layout)


    def have_internet_connection(self):
        try:
            socket.create_connection(("www.codal.ir", 80))
            return True
        except OSError:
            pass
        return False


    def doCrawling(self):
        self._crawl(None, CodalSpider)
        self.process.start()
        


    def sleep(self, *args, seconds):
        """Non blocking sleep callback"""
        return deferLater(reactor, seconds, lambda: None)

    def _crawl(self,result,spider):
        deferred = self.process.crawl(spider)
        deferred.addCallback(lambda results: print('waiting 2 seconds before restart...'))
        if(not self.stopFlag):
            deferred.addCallback(self.sleep, seconds=2)
            deferred.addCallback(self._crawl, spider)
        else:
            while True:
                # print("idle Loop")
                a=1 # do nothing
                if(self.exitFlag):
                    break
                if(not self.stopFlag):
                    deferred.addCallback(self._crawl, spider)
                    break
        return deferred

    def buttonGoClicked(self):
        if(self.have_internet_connection()):
            self.addedData.clear()
            while (self.tableWidget.rowCount() > 1):
                self.tableWidget.removeRow(1)
            self.stopFlag=False
            self.spinner.start()
            self.input_symbols_text.clear()
            for textInputObj in self.symbols_textInput_objs:
                self.input_symbols_text.append(textInputObj.text())
            self.input_keywords_text.clear()
            for keywordInputObj in self.keywords_textInput_objs:
                self.input_keywords_text.append(keywordInputObj.text())
            if(not self.secondRunFlag): #first run
                threading.Timer(1.0, self.doCrawling).start()
            else:
                self.stopFlag=False
            self.secondRunFlag=True
        else:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Critical)
            msg.setText("No internet Connection!")
            msg.setInformativeText('')
            msg.setWindowTitle("Error")
            msg.exec_()


    def buttonStopClicked(self):
        self.stopFlag=True
        self.spinner.stop()

    def buttonExitClicked(self):
        self.stopFlag=True
        self.exitFlag=True
        time.sleep(1)
        sys.exit()

    def buttonClearClicked(self):
        while (self.tableWidget.rowCount() > 1):
            self.tableWidget.removeRow(1)


    def addDataToTableWidget(self,symbol,cname,title,date):
        if(self.selective_radioButton.isChecked()):
            if symbol==self.input_symbols_text[0] or symbol==self.input_symbols_text[1] or symbol==self.input_symbols_text[2] or symbol==self.input_symbols_text[3] or symbol==self.input_symbols_text[4] or symbol==self.input_symbols_text[5] or symbol==self.input_symbols_text[6] or symbol==self.input_symbols_text[7] or symbol==self.input_symbols_text[8] or symbol==self.input_symbols_text[9] and symbol!='':
                conditionsflag=False
                if(self.with_keyword_checkbox.isChecked()):
                    if self.input_keywords_text[0] in title:
                        if self.input_keywords_text[0] != '':
                            conditionsflag=True
                    elif self.input_keywords_text[1] in title:
                        if self.input_keywords_text[1] != '':
                            conditionsflag=True
                    elif self.input_keywords_text[2] in title:
                        if self.input_keywords_text[2] != '':
                            conditionsflag=True
                    elif self.input_keywords_text[3] in title:
                        if self.input_keywords_text[3] != '':
                            conditionsflag=True
                    elif self.input_keywords_text[4] in title:
                        if self.input_keywords_text[4] != '':
                            conditionsflag=True
                else:
                    # without keyword mode
                    conditionsflag=True


                #Check Is Data Duplicate ?
                duplicateFlag=False
                if conditionsflag:
                    for data in self.addedData:
                        if data[0]==symbol and data[1]==title and data[2]==date:
                            duplicateFlag=True


                if conditionsflag and not duplicateFlag:
                    row_number = 1
                    self.tableWidget.insertRow(row_number)
                    self.tableWidget.setItem(row_number,3, QTableWidgetItem(symbol))
                    self.tableWidget.setItem(row_number,2, QTableWidgetItem(cname))
                    self.tableWidget.setItem(row_number,1, QTableWidgetItem(title))
                    self.tableWidget.setItem(row_number,0, QTableWidgetItem(date))

                    item = self.tableWidget.item(row_number, 3)
                    item.setTextAlignment(Qt.AlignCenter)
                    item = self.tableWidget.item(row_number, 2)
                    item.setTextAlignment(Qt.AlignCenter)
                    item = self.tableWidget.item(row_number, 1)
                    item.setTextAlignment(Qt.AlignCenter)
                    item = self.tableWidget.item(row_number, 0)
                    item.setTextAlignment(Qt.AlignCenter)

                    winsound.MessageBeep(winsound.MB_ICONHAND) #play sound on add new
                    self.addedData.append([symbol,title,date])


        
        elif(self.all_radioButton.isChecked()):
            conditionsflag=False
            if(self.with_keyword_checkbox.isChecked()):
                if self.input_keywords_text[0] in title:
                    if self.input_keywords_text[0] != '':
                        conditionsflag=True
                elif self.input_keywords_text[1] in title:
                    if self.input_keywords_text[1] != '':
                        conditionsflag=True
                elif self.input_keywords_text[2] in title:
                    if self.input_keywords_text[2] != '':
                        conditionsflag=True
                elif self.input_keywords_text[3] in title:
                    if self.input_keywords_text[3] != '':
                        conditionsflag=True
                elif self.input_keywords_text[4] in title:
                    if self.input_keywords_text[4] != '':
                        conditionsflag=True
            else:
                # without keyword mode
                conditionsflag=True


            #Check Is Data Duplicate ?
            duplicateFlag=False
            if conditionsflag:
                for data in self.addedData:
                    if data[0]==symbol and data[1]==title and data[2]==date:
                        duplicateFlag=True


            if conditionsflag and not duplicateFlag:
                row_number = 1
                self.tableWidget.insertRow(row_number)
                self.tableWidget.setItem(row_number,3, QTableWidgetItem(symbol))
                self.tableWidget.setItem(row_number,2, QTableWidgetItem(cname))
                self.tableWidget.setItem(row_number,1, QTableWidgetItem(title))
                self.tableWidget.setItem(row_number,0, QTableWidgetItem(date))

                item = self.tableWidget.item(row_number, 3)
                item.setTextAlignment(Qt.AlignCenter)
                item = self.tableWidget.item(row_number, 2)
                item.setTextAlignment(Qt.AlignCenter)
                item = self.tableWidget.item(row_number, 1)
                item.setTextAlignment(Qt.AlignCenter)
                item = self.tableWidget.item(row_number, 0)
                item.setTextAlignment(Qt.AlignCenter)

                winsound.MessageBeep(winsound.MB_ICONHAND) #play sound on add new
                self.addedData.append([symbol,title,date])
        

if __name__ == '__main__':
    appctxt = ApplicationContext()
    gui = Widget_GUI()
    gui.show()
    sys.exit(appctxt.app.exec_())
