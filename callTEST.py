import sys , sqlite3
from PyQt5.QtWidgets import QDialog, QApplication, QComboBox, QMessageBox
from PyQt5.QtWidgets import QApplication, QMainWindow,  QAction, QTextEdit, QFontDialog, QColorDialog
from sqlite3 import Error
from PyQt5.QtPrintSupport import QPrinter, QPrintPreviewDialog
from PyQt5 import QtGui, QtCore, QtWidgets
import time
#********************CHART***********************
import matplotlib.pyplot as plt
#******************** Export to excel ***********************
import xlsxwriter
from xlsxwriter.workbook import Workbook
workbook = xlsxwriter.Workbook('عمولات.Xlsx')
worksheet = workbook.add_worksheet()
#******
from TEST import *
class MyForm (QDialog):
    def __init__(self):
        super() .__init__()
        self.ui = Ui_Dialog()
        self.ui.setupUi(self)
        self.show()
        self.ui.ente.clicked.connect(self.CalculateComm)
        self.ui.DTS.clicked.connect(self.ADDDTS)
        self.ui.PRNTT.clicked.connect(self.PENT)
        self.ui.EDT.clicked.connect(self.EDIT)
        self.ui.CLS.clicked.connect(self.SAVE)
        self.ui.FNT.clicked.connect(self.FONT)
        self.ui.nw.clicked.connect(self.NEWITEM)
        self.ui.crt.clicked.connect(self.CREATEDB)
        self.ui.DLT.clicked.connect(self.DELETE)
        self.ui.exprt.clicked.connect(self.CHART)
        self.ui.Excel.clicked.connect(self.EXLS)
    def CREATEDB(self):
        conn = sqlite3.connect("ALLCOM.db")
        c = conn.cursor()
        print("Database is opened succefully")
        c.execute('CREATE TABLE COMM ( ID INT PRIMARY KEY  NOT NULL ,unit data char(30),Unit no char(30), deposit INT, rent INT , period time char(30),due rent int , paid rent int , commision int )')
        print("table is created")
        c.close()
    def CalculateComm(self):
        try:
           RR = int(self.ui.rent.text())
           TT = int(self.ui.taa.text())
           RE = str(self.ui.rente.text())
           NET = str(self.ui.nete.text())
           COMM = str(self.ui.comm.text())
           DAW = str(self.ui.daw.currentText())
        except Exception as er:
            print(er)
            QMessageBox.information(self.ui.ente,"خطاء", "من فضلك ادخل الحقول الفارغة أولاً  ً")
            return
        if DAW == "":
            QMessageBox.information(self.ui.ente ,"ملاحظة", "من فضلك إختر دورة السداد من القائمة  ")
        if DAW == "شهري":
            RE =(int(RR) / 12 + int(TT))
            self.ui.rente.setText((str(RE)))
            COMM = int(RE) - int(NET)
            self.ui.comm.setText((str(COMM)))
        elif DAW == "ثلث سنوي":
            RE = (int(RR) / 3 + int(TT))
            self.ui.rente.setText((str(RE)))
            COMM = int(RE) - int(NET)
            self.ui.comm.setText((str(COMM)))
        elif DAW == "سدس سنوي":
            RE = (int(RR) / 6 + int(TT))
            self.ui.rente.setText((str(RE)))
            COMM = int(RE) - int(NET)
            self.ui.comm.setText((str(COMM)))
        elif DAW == "ربع سنوي":
            RE = (int(RR) / 4 + int(TT))
            self.ui.rente.setText((str(RE)))
            COMM = int(RE) - int(NET)
            self.ui.comm.setText((str(COMM)))
        elif DAW == "نصف سنوي":
            RE = (int(RR) / 2 + int(TT))
            self.ui.rente.setText((str(RE)))
            COMM = int(RE) - int(NET)
            self.ui.comm.setText((str(COMM)))
        elif DAW == "سنوي":
            RE = (int(RR) / 1 + int(TT))
            self.ui.rente.setText((str(RE)))
            COMM = int(RE) - int(NET)
            self.ui.comm.setText((str(COMM)))

    # ********************************ADD NEW ITEM ***********************************
    def NEWITEM(self):
        m2 = QMessageBox.question(self.ui.nw,"سؤال","هل تريد اضافة بيانات جديدة",
                                  QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
        if m2 == QMessageBox.Yes:
            self.ui.em.clear()
            self.ui.rqm.clear()
            self.ui.nete.clear()
            self.ui.rente.clear()
            self.ui.rent.clear()
            self.ui.taa.clear()
            self.ui.comm.clear()
            self.ui.crt.hide()
        if m2 == QMessageBox.No:
            #self.ui.crt.hide()
            return
    # ********************************ADD DATA TO WIDGETS***********************************
    def ADDDTS(self):
        conn = sqlite3.connect('ALLCOM.db')
        c = conn.cursor()
        SQL = "INSERT INTO COMM VALUES ( '%s' , '%s' , '%s' ,'%s', '%s' , '%s' ,'%s' ,'%s') " % (
        self.ui.em.text() , self.ui.rqm.text() , self.ui.taa.text() , self.ui.rent.text() , self.ui.daw.currentText() ,
        self.ui.rente.text() , self.ui.nete.text() , self.ui.comm.text())
        try:
            c.execute(SQL)
            self.ui.VIEW.append(self.ui.label_3.text())
            self.ui.VIEW.append(self.ui.em.text())
            self.ui.VIEW.append(self.ui.label_2.text())
            self.ui.VIEW.append(self.ui.rqm.text())
            self.ui.VIEW.append(self.ui.label_5.text())
            self.ui.VIEW.append(self.ui.taa.text())
            self.ui.VIEW.append(self.ui.label_6.text())
            self.ui.VIEW.append(self.ui.rent.text())
            self.ui.VIEW.append(self.ui.label_7.text())
            self.ui.VIEW.append(self.ui.daw.currentText())
            self.ui.VIEW.append(self.ui.label.text())
            self.ui.VIEW.append(self.ui.nete.text())
            self.ui.VIEW.append(self.ui.label_8.text())
            self.ui.VIEW.append(
                self.ui.comm.text() + "\n                ===================================================")
            conn.commit()
        except Error as ef:
            conn.rollback()
            p = print(ef)
            QMessageBox.information(self.ui.DTS,"معلومة",str(p))

        conn.close()
# ********************************معاينة الطباعة***********************************
    def PENT(self):
        Namee = self.ui.rqm.text()
        conn = sqlite3.connect('ALLCOM.db')
        c = conn.cursor()
        try:
            SQLl = c.execute("SELECT sum(commision) FROM COMM WHERE unit like ?", (Namee,))
            for i in SQLl:
                print(i)

            #   conn.commit()
            # for row in SQLl:
            # print(str(row[5]))
            # t1 = row[5] * -1
            self.ui.VIEW.append("إجمالي ايرادات المندوب:" + Namee + str(i) + " ريال سعودي")
            self.ui.VIEW.append("the total commision of sals rep is: ".title() + str(i) + " sar" + Namee)
        except Exception as eee:
            print(eee)
        c.close()
        printer = QPrinter(QPrinter.HighResolution)
        previewDialog = QPrintPreviewDialog(printer, self)
        previewDialog.paintRequested.connect(self.printPreview)
        previewDialog.exec_()
    def printPreview(self, printer):
        self.ui.VIEW.print_(printer)

# *******************************EDIT DATAS***********************************
    def EDIT(self):
        m1 = QMessageBox.question(self.ui.EDT,"سؤال","هل تريد تعديل البيانات المُدخلة",
                                  QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
        if m1 == QMessageBox.Yes:
            try:
                conn = sqlite3.connect('ALLCOM.db')
                c = conn.cursor()
                emar = str(self.ui.em.text())
                #rentt = self.ui.rent.text()
                SQL1 = ('SELECT * FROM COMM  WHERE ID = $s')
                c.execute(SQL1, [emar])
                data = c.fetchone()
                print(data)
                self.ui.rqm.setText(data[1])
                self.ui.taa.setText(str(data[2]))
                self.ui.rent.setText(str(data[3]))
                self.ui.daw.setCurrentText(str(data[4]))
                self.ui.rente.setText(str(data[5]))
                self.ui.nete.setText(str(data[6]))
                self.ui.comm.setText(str(data[7]))
            except Exception as ee:
                print(ee)

           # c.execute("UPDATE COMM SET unit = 'العدروسي' WHERE ID = (?)",(35,))
            #print("تم تعديل البيانات")
            #conn.commit()
            #conn.close()
            self.ui.VIEW.setReadOnly(False)
        else:
            self.ui.VIEW.setReadOnly(True)
# ********************************SAVE DATA AFTER EDIT ***********************************
    def SAVE(self):
        try:
            raqm = self.ui.rqm.text()
            emaraa = self.ui.em.text()
            DEOST = self.ui.taa.text()
            COMISION = self.ui.comm.text()
            DAWRA = self.ui.daw.currentText()
            conn = sqlite3.connect('ALLCOM.db')
            c = conn.cursor()
            SQL1 = 'UPDATE COMM SET unit = ?, deposit = ? , period = ? , commision = ?  WHERE ID = $s'
            VAL = (raqm  , DEOST, DAWRA , COMISION,  emaraa)
            c.execute(SQL1 , VAL)
            conn.commit()
            self.ui.VIEW.append("تم تعديل البيانات")
            QMessageBox.information(self.ui.CLS,"معلومة","تم تعديل البيان")
            M3 = QMessageBox.question(self.ui.CLS,"سؤال","هل تريد عرض كود المندوب؟ YES تعني عرض رقم المندوب / NO تعني عرض اجمالي ايرادات لكل مندوب ",
                                      QMessageBox.No| QMessageBox.Yes,QMessageBox.Yes)
            if M3 == QMessageBox.Yes:
                try:
                    conn = sqlite3.connect('ALLCOM.db')
                    c = conn.cursor()
                    SQL4 = 'SELECT ID , unit From COMM '
                    t1 = c.execute(SQL4)
                    for row in t1:
                        print(row)
                       # print(row[1])
                        self.ui.VIEW.append(str(row))
                    time.sleep(2)
                except Exception as eed:
                    print(eed)
            if M3 == QMessageBox.No:
                try:
                    conn = sqlite3.connect('ALLCOM.db.')
                    c = conn.cursor()
                    SQL3 = 'SELECT unit, SUM (commision) FROM COMM GROUP BY unit '
                    t2 = c.execute(SQL3)
                    for row in t2:
                        print(row)
                        self.ui.VIEW.append(str(row))
                    time.sleep(2)
                except Exception as ef:
                    print(ef)
                #self.ui.VIEW.append(t1)
            else:
                return
           # conn.close()
        except Error as ee:
            print(ee)
#       try:
#           m2 = QMessageBox.question(self.ui.CLS,"سؤال", "هل تريد إغلاق البرنامج",
     #                                QMessageBox.Yes|QMessageBox.No, QMessageBox.Yes)
 #          if m2 == QMessageBox.Yes:
  #            w.close()
   #        if m2 == QMessageBox.No:
     #          return
    #   except Exception as ee:
      #    print(ee)"""

# ******************************** تعديل الخط  ***********************************
    def FONT(self):
        font, ok = QFontDialog.getFont()
        if ok:
            self.ui.VIEW.setFont(font)
            time.sleep(1)
        #conn = sqlite3.connect('ALLCOM.db')
        #c = conn.cursor()
        #c.execute("DELETE FROM COMM WHERE ID = 10 ")
        #conn.commit()
        #conn.close()
# ******************************** delete data  ***********************************
    def DELETE(self):
        conn = sqlite3.connect('ALLCOM.db')
        c = conn.cursor()
        emara = int(self.ui.em.text())
        listt = range(1 , 5000 )
        if emara in listt :
            c.execute("DELETE FROM COMM WHERE ID like(?)",(emara,))
            self.ui.em.clear()
            self.ui.rqm.clear()
            self.ui.nete.clear()
            self.ui.rente.clear()
            self.ui.rent.clear()
            self.ui.taa.clear()
            self.ui.comm.clear()
            self.ui.VIEW.setText("تم مسح البيانات          Data is deleted")
            conn.commit()
            conn.close()
            QMessageBox.information(self.ui.DLT,"معلومة","تم مسح البيان")
        else:
            QMessageBox.information(self.ui.DLT,"معلومة","الرقم غير موجود")
 # *********************** DO CHARTS ***********
    def CHART(self):
        try:
            conn = sqlite3.connect('ALLCOM.db')
            c = conn.cursor()
            SQL3 = c.execute("SELECT unit, SUM(commision) FROM COMM GROUP BY unit")
            for row in SQL3:
                n1 = [(str(row[0]))]
                r1 = [(int(row[1]*-1))]
                plt.bar(n1 , r1)
            plt.title("revenue of markters".title())
            plt.xlabel("names of markter".title())
            plt.ylabel("revenue".title())
           # plt.scatter(data_x, data_y, s=size, c=color, alpha=.5)
            plt.grid(True)
            plt.show()
            conn.commit()
            c.close()
        except EXCEPTION as ed:
            print(ed)
 # *********************** EXPORT TO XLSX***********
    def EXLS (self):
        try:
            conn = sqlite3.connect('ALLCOM.db')
            c = conn.cursor()
            SQL4 = c.execute('SELECT * FROM COMM')
            for p, row in enumerate(SQL4):
                for q, value in enumerate(row):
                    worksheet.write(p, q, value)
            workbook.close()
            self.ui.VIEW.setText("تم تصدير البيانات")
        except Error as ed:
            print(ed)
if __name__=="__main__":
    app = QApplication(sys.argv)
    w = MyForm()
    w.show()
    sys.exit(app.exec_())
input()