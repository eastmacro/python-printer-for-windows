import sys
import os
import json
import re
import xlrd
import win32ui
import win32gui
import win32print
import win32con

import config

import PyQt4.QtNetwork
from PyQt4.QtCore import *
from PyQt4.QtGui import *
from PyQt4.QtWebKit import *
# -*- coding: utf-8 -*-

class WebViewSupportingEsc(QWebView):
  def keyPressEvent(self, event):
    if event.key() == Qt.Key_Escape:
      self.close()
    return super().keyPressEvent(event)

class Printer(QObject):
  def __init__(self, parent=None):
    super().__init__(parent)
    self.sourcePath = None
    self.index = 1
    self.total = 1

  def _do_changeNum(self, text=None):
    adict = {
      '1':'一',
      '2':'二',
      '3':'三',
      '4':'四',
      '5':'五',
      '6':'六',
      '7':'七',
      '8':'八',
      '9':'九',
      '0':'○'
    }
    rx = re.compile('|'.join(map(re.escape, adict)))  
    def one_xlat(match):  
      return adict[match.group(0)]  
    return rx.sub(one_xlat, str(text)) 
    
  @pyqtSlot(str)
  def do_getPrinters(self, args=None):
    global frame
    _printers = []
    printers = {}
    try:
      printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL, None, 4)
    except:
      frame.evaluateJavaScript('alert("枚举打印机失败-rpc服务不可用还是打印机服务没启动");');
    
    if printers:
      for _printer in printers:
        _printers.append(_printer['pPrinterName'])
    _printers = json.dumps(_printers)
    frame.evaluateJavaScript('var printers=\'%s\';null' % (_printers,));

  @pyqtSlot(str)
  def do_getNextData(self, num=0):
    if self.index <= self.total:
      if num == '1':
        self.index  += 1
      else:
        self.index  -= 1
      if self.index < 1:
        self.index = 1
      if self.index > self.total:
        self.index = self.total
    else:
      self.index = self.total
      return

    row_data = self.do_getOneData(self.index)
    frame.evaluateJavaScript('var data_user=\'%s\'; var index=%d; null' % (json.dumps(row_data),self.index));
    
  @pyqtSlot(str)
  def do_getOneData(self, num=0):
    bk = xlrd.open_workbook(self.sourcePath)
    sh = bk.sheet_by_index(0)
    if num > (sh.nrows-1):
      return 
    else:
      row_data = sh.row_values(num)
      return row_data

  @pyqtSlot(str)
  def do_setSourceFullPath(self, file=''):
    dlg = win32ui.CreateFileDialog(1)
    dlg.DoModal()
    filename = dlg.GetPathName()
    if filename == '':
      return
    self.sourcePath = filename
    bk = xlrd.open_workbook(self.sourcePath)
    try:
      sh = bk.sheet_by_index(0)
    except:
      message="格式错误"
    self.total = sh.nrows-1
    message=" /%d人 " % (self.total,)
    data_head = json.dumps(self.do_getOneData(0))
    data_user = json.dumps(self.do_getOneData(1))

    frame.evaluateJavaScript('var sourceFullPath=\'%s\';var message="%s";var data_head=\'%s\';var data_user=\'%s\';null' % (filename.replace('\\', '\\\\'),message, data_head, data_user));

  @pyqtSlot(str)
  def do_print(self, print_name=''):
    row_data = self.do_getOneData(self.index)
    #print(row_data)
    pHandle = win32print.OpenPrinter(print_name)
    printinfo = win32print.GetPrinter(pHandle,2)
    pDevModeObj = printinfo["pDevMode"]                           
    pDevModeObj.Orientation = win32con.DMORIENT_LANDSCAPE 
    pDevModeObj.PaperSize   = 9 #8

    _dc =win32gui.CreateDC('WINSPOOL',print_name,pDevModeObj)
    dc=win32ui.CreateDCFromHandle(_dc)
    dc.SetMapMode(win32con.MM_TWIPS) # 1440 per inch
    scale_factor = 14.2 # i.e. 14 twips to the point  
    scale_factor_height = scale_factor # i.e. 14 twips to the point  
    dc.StartDoc(row_data[1])
    dc.StartPage()

    pen = win32ui.CreatePen(0, int(scale_factor), 0)
    dc.SelectObject(pen)
    font = win32ui.CreateFont({
        "name": "方正仿宋_GBK",
        "height": int(scale_factor * 24),
        "weight": 400,
    })
    dc.SelectObject(font)

    dc.TextOut(int(428*scale_factor), int(-1 * scale_factor_height * config.liney_1), row_data[2]) #姓名
    dc.TextOut(int(650*scale_factor),int(-1 * scale_factor_height * config.liney_1), row_data[3])  #性别
    dc.TextOut(int(836*scale_factor),int(-1 * scale_factor_height * config.liney_1), row_data[4])  #身份

    start_year = self._do_changeNum(str(int(row_data[7]))[0:4]);
    start_month = self._do_changeNum(str(int(row_data[7]))[5:]);
    dc.TextOut(int(361*scale_factor),int(-1 * scale_factor_height * config.liney_2), start_year)  #开始日期
    dc.TextOut(int(539*scale_factor),int(-1 * scale_factor_height * config.liney_2), start_month)  #开始日期
    end_year = self._do_changeNum(str(int(row_data[8]))[0:4]);
    end_month = self._do_changeNum(str(int(row_data[8]))[5:]);
    dc.TextOut(int(660*scale_factor),int(-1 * scale_factor_height * config.liney_2), end_year)  #结束日期
    dc.TextOut(int(824*scale_factor),int(-1 * scale_factor_height * config.liney_2), end_month)  #结束日期
    year = self._do_changeNum(str(int(row_data[9])));
    dc.TextOut(int(1004*scale_factor),int(-1 * scale_factor_height * config.liney_2), year)  #学制

    dc.TextOut(int(459*scale_factor),int(-1 * scale_factor_height * config.liney_3), row_data[6])  #专业
    dc.TextOut(int(734*scale_factor),int(-1 * scale_factor_height * config.liney_3), config.graduation_type)  #学历
    dc.TextOut(int(450*scale_factor),int(-1 * scale_factor_height * config.liney_4), row_data[1])  #证书编号
    dc.TextOut(int(805*scale_factor),int(-1 * scale_factor_height * config.liney_4), config.graduation_time_year)  #毕业时间
    dc.TextOut(int(950*scale_factor),int(-1 * scale_factor_height * config.liney_4), config.graduation_time_month)  #
    dc.TextOut(int(1022*scale_factor),int(-1 * scale_factor_height * config.liney_4), config.graduation_time_dat)  #

    dc.EndPage()
    dc.EndDoc()

def main():
  global frame, index
  index = 0
  app = QApplication(sys.argv)
  web = WebViewSupportingEsc()
  web.setWindowTitle(u'毕业证批量打印')
  web.setFixedSize(800, 500)

  web.setContextMenuPolicy(Qt.NoContextMenu)
  frame = web.page().mainFrame()
  printerObj = Printer(frame)

  frame.addToJavaScriptWindowObject('printerObj', printerObj)
  frame.javaScriptWindowObjectCleared.connect(
    lambda: frame.addToJavaScriptWindowObject('printerObj', printerObj))

  web.load(QUrl('./resource/main.html'))

  web.show()
  sys.exit(app.exec_())

if __name__ == '__main__':
  main()


