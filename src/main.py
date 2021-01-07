"""
	These lines are for needed librarys. gui is not a library but it includes the main app class(its splited
	because it would make this code more messy then it is).
	Sqlite is the database for storing data.
	XlsxWriter is for writing xlsx files.
	Xlrd is for reading xlsx files.
	PyQt is the base gui library.
	Sys is one of useful python standerd library.
"""
import gui
import sqlite3 as db
import xlsxwriter
import xlrd
import functools
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QApplication, QMessageBox, QFileDialog
from PyQt5.QtCore import pyqtSlot
from sys import argv


"""
	This class is the main app which extends the empty program window to a working one.
"""
class App(QtWidgets.QMainWindow, gui.Ui_MainWindow):
	
	# This function runs first for adding elements
    def __init__(self, parent=None):
		# We init and add the ui elements to main window
        super(App, self).__init__(parent)
        self.setupUi(self)
        # Load the datatable with a full quary
        self.load_table('SELECT * FROM phones')
        # Maximize the window
        self.showMaximized()

	
	# This function updates the changed data in the datatable
    @functools.cache
    def save(self):
		# We loop in rows and cols
        cur = con.cursor()
        for i in range(self.table.rowCount()):
			# This list resets every row for getting the new data
            column = []
            for j in range(self.table.columnCount()):
				# We skip the 12 col because its a combobox and cant be saved with the normal method
                if j != 12:
                    column.append(self.table.model().data(self.table.model().index(i, j)))
                    # This returns the data of table and we can extract the data by using data method
                    # self.table.model().data()
                    # This returns value of the normal textbox:
                    # self.table.model().index(i, j)
                    # And we combine them and append them to the list
                else:
					# We append the value of selected combobox and append it to column list for later
                    column.append(self.table.cellWidget(i,j).currentText())
            
            # Here we make a standerd update sql quary with the values we got
            sql = """UPDATE Phones 
            SET name = '%s',
            family = '%s',
            phone1 = '%s',
            phone2 = '%s',
            phone3 = '%s',
            home1 = '%s',
            home2 = '%s',
            work_number = '%s',
            home_path = '%s',
            fax = '%s',
            website = '%s',
            email = '%s',
            messager = '%s',
            phone_msg = '%s',
            workpath = '%s'
            WHERE id = %s;""" % tuple(column)
            # We execute the quary
            cur.execute(sql)
            # And commit it to be sure
            con.commit()
		# The code above runs for all rows
 
    def clear_table(self):
		# By setting the row count to 0 we can reset the datatable
        self.table.setRowCount(0)
	
	# This function loads the datatable with a givien quary
    def load_table(self,sql):
		# We get the data from database functions
        output = LoadData(sql)
        # If the output was false we return false which means user not found
        if output == []:
            return False
        
        for row in output:
			# We get the postion of datatable row
            row_pos = self.table.rowCount()
            # This inserts a blank row
            self.table.insertRow(row_pos)
            # This loop fills the row with data
            for i, column in enumerate(row, 0):
				# This flags checks if the current item is a combobox or a normal item
                normal_widget_item = True
                # This creates a needed item for inserting to the row
                item = QtWidgets.QTableWidgetItem(str(column))
                
                # The 12 col is always a combobox
                if i == 12:
					# We create a combobox like the combobox like in the main window
                    item = QtWidgets.QComboBox()
					# This addeds all options to select
                    item.addItems(self.all_messager_types)
                    # Sets the default item in combobox
                    item.setCurrentIndex(self.all_messager_types.index(column))
                    # Set the flag so it can change the way to insert
                    normal_widget_item = False
				
				# The 15 one is the id part which is very important to not change
                if i == 15:
					# This sets the item status to be anything but not editable
                    flags = QtCore.Qt.ItemFlags()
                    flags != QtCore.Qt.ItemIsEnabled
                    item.setFlags(flags)
                
                # The normal way of adding the item
                if normal_widget_item:
                    self.table.setItem(row_pos, i, item)
                # Adding the widget item
                else:
                    self.table.setCellWidget(row_pos,i,item)
        
        #This resizes cols to be neat and good
        self.table.resizeColumnsToContents()
	
	# Dialog box for errors
    def error(self, text, title = "مشکل"):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Critical)
        msg.setText(text)
        msg.setWindowTitle("مشکل")
        msg.exec_()
        del msg
    # Dialog box for infos
    def info(self, text, title = "اطلاعات"):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setText(text)
        msg.setWindowTitle("اطالاعات")
        msg.exec_()
        del msg
	# Dialog box for asking a question
    def question(self,title,text):
        buttonReply = QMessageBox.question(self, title, text, QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if buttonReply == QMessageBox.Yes:
            del buttonReply
            return True
        else:
            del buttonReply
            return False
    
    # Exporting a excel file for backup or etc
    @functools.cache
    def export_excel(self):
		# A dialog box for saving a file
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getSaveFileName(self,"Export", "","Xls Files (*.xlsx)", options=options)
        # This checks if user saved a file
        if fileName:
			# This add .xlsx to end of file to make sure MS Windows os supports
            fileName += '.xlsx'
            # Creating a workbook and sheet for saving the data (At the selected path)
            workbook = xlsxwriter.Workbook(fileName)
            worksheet = workbook.add_worksheet()
            
            # Adding the titles to excel file
            for i in range(self.table.columnCount()):
					# Getting the value of every header(title)
                    text = self.table.horizontalHeaderItem(i).text()
                    # Writing the value at row 0 and correct col
                    worksheet.write(0, i,text)
			
			# Adding the datas to excel file
            for i in range(self.table.columnCount()):
                for j in range(self.table.rowCount()):
					# The 12 col is speacial and we skip it
                    if i != 12:
						# Getting the value of normal cols
                        text = self.table.item(j, i).text()
                    else:
						# Getting the value of selected combobox
                        text = self.table.cellWidget(j,i).currentText()
                    # Writing the value into excel file
                    worksheet.write(j + 1, i,text)
			# Closing and saving the file
            workbook.close()
            # 🤩︎
            self.info('خروجی با موفقیت ایجاد شد !')
    
    @functools.cache
    def import_excel(self):
        options = QFileDialog.Options()
        options = QFileDialog.DontUseNativeDialog
        user_file, _ = QFileDialog.getOpenFileName(self,"Please select a file", "","Excel File (*.xlsx)", options=options)
        if user_file:
            wb = xlrd.open_workbook(user_file)
            sheet = wb.sheet_by_index(0)
            
            for i,row in enumerate(range(sheet.nrows)):
                if i == 0:
                    continue
                row_data = []
                for j,col in enumerate(range(sheet.ncols)):
                    if j == 15:
                        continue
                    row_data.append(sheet.cell_value(row, col))
                try:
                    AddData(row_data)
                    self.load_table('SELECT * FROM Phones')
                    self.info('اطلاعات با موفقیت وارد شد !')
                
                except db.IntegrityError:
                    self.error('کاربر شماره {0} موجود است!'.format(i))
                
                except db.ProgrammingError:
                    self.error('فایل اکسل دستکاری شده است. کاربر {0} به پایگاه داده اضافه نمیشود.'.format(i))

    def quit_safe(self):
        self.save()
        exit(0)
    
    def closeEvent(self, event):
        self.quit_safe()
        event.accept()

    def about_programer(self):
        self.info("این برنامه توسط علیرضا آهنی ساخته شده است.\nبرای تماس میتونید از ایمیل alirezaahani@protonmail.com\nاستفاده کنید")

    def about_licenc(self):
        self.info('این برنامه تحت مجوز نرم افزاری GPL نسخه ی ۳ منتشر شده است.')
    
    def reset_textboxs(self):
        self.clear_table()
        self.load_table('SELECT * FROM phones')
        self.name.setText("")
        self.family.setText("")
        self.phone1.setText("")
        self.phone2.setText("")
        self.phone3.setText("")
        self.home1.setText("")
        self.home2.setText("")
        self.work_number.setText("")
        self.home_path.setText("")
        self.fax.setText("")
        self.website.setText("")
        self.email.setText("")
        self.phone_msg.setText("")
        self.workpath.setText("")
    
    @pyqtSlot()
    @functools.cache
    def add_button(self):
        datas = {
            'name' : self.name.text(),
            'family' : self.family.text(),
            'phone1' : self.phone1.text(),
            'phone2' : self.phone2.text(),
            'phone3' : self.phone3.text(),
            'home1' : self.home1.text(),
            'home2' : self.home2.text(),
            'work_number' : self.work_number.text(),
            'home_path' : self.home_path.text(),
            'fax' : self.fax.text(),
            'website' : self.website.text(),
            'email' : self.email.text(),
            'messager' : self.messager.currentText(),
            'phone_msg' : self.phone_msg.text(),
            'workpath' : self.workpath.text()
        }
        if datas['name'] and datas['family']:
            try:
                AddData(list(datas.values()))
                self.reset_textboxs()
            except db.IntegrityError:
                self.error('اطالاعات مورد نظر در پایگاه داده موجود است')
        else:
            self.error('لطفا نام و نام خانوادگی را پرکنید')
    
    @pyqtSlot()
    @functools.cache
    def search_button(self):
        datas = {
            'name' : self.name.text(),
            'family' : self.family.text(),
            'phone1' : self.phone1.text(),
            'phone2' : self.phone2.text(),
            'phone3' : self.phone3.text(),
            'home1' : self.home1.text(),
            'home2' : self.home2.text(),
            'work_number' : self.work_number.text(),
            'home_path' : self.home_path.text(),
            'fax' : self.fax.text(),
            'website' : self.website.text(),
            'email' : self.email.text(),
            'messager' : self.messager.currentText(),
            'phone_msg' : self.phone_msg.text(),
            'workpath' : self.workpath.text()
        }
        sql = """
        SELECT * FROM Phones WHERE 
        name LIKE '%{0}%' AND 
        family LIKE '%{1}%' AND 
        phone1 LIKE '%{2}%' AND 
        phone2 LIKE '%{3}%' AND
        phone3 LIKE '%{4}%' AND
        home1 LIKE '%{5}%' AND
        home2 LIKE '%{6}%' AND
        work_number LIKE '%{7}%' AND
        home_path LIKE '%{8}%' AND
        fax LIKE '%{9}%' AND
        website LIKE '%{10}%' AND
        email LIKE '%{11}%' AND
        messager LIKE '%{12}%' AND
        phone_msg LIKE '%{13}%' AND
        workpath LIKE '%{14}%'
        """.format(
            datas['name'],
            datas['family'],
            datas['phone1'],
            datas['phone2'],
            datas['phone3'],
            datas['home1'],
            datas['home2'],
            datas['work_number'],
            datas['home_path'],
            datas['fax'],
            datas['website'],
            datas['email'],
            datas['messager'],
            datas['phone_msg'],
            datas['workpath']
        )
        self.clear_table()
        if self.load_table(sql) == False:
            self.error("کاربری پیدا نشد.")
    
    @pyqtSlot()
    def delete_button(self):
        if self.table.selectionModel().selectedRows() == []:
            self.error("لطفا کاربری را برای حذف انتخاب کنید")
            return False
        
        if self.question("حذف کاربر","آیا مطمئن هستید؟"):
            for index in sorted(self.table.selectionModel().selectedRows()):
                row = index.row()
                sql = "DELETE FROM Phones WHERE name LIKE '%{0}%' AND family LIKE '%{1}%' AND phone1 LIKE '%{2}%' AND id = '{3}'".format(
                    self.table.model().data(self.table.model().index(row, 0)),
                    self.table.model().data(self.table.model().index(row, 1)),
                    self.table.model().data(self.table.model().index(row, 2)),
                    self.table.model().data(self.table.model().index(row, 15))
                )
                cur = con.cursor()
                cur.execute(sql)
            con.commit()
            self.clear_table()
            self.load_table('SELECT * FROM Phones')
    
    @pyqtSlot()
    def export(self):
        self.export_excel()

def CreateTable():
    global con
    cur = con.cursor()
    cur.execute('''
        CREATE TABLE IF NOT EXISTS `phones`(
            name TEXT,
            family TEXT,
            phone1 TEXT,
            phone2 TEXT,
            phone3 TEXT,
            home1 TEXT,
            home2 TEXT,
            work_number TEXT,
            home_path TEXT,
            fax TEXT,
            website TEXT,
            email TEXT,
            messager TEXT,
            phone_msg TEXT,
            workpath TEXT,
            id INTEGER PRIMARY KEY AUTOINCREMENT
        )
        ''')
    return True

@functools.cache
def LoadData(sql):
    cur = con.cursor()
    cur.execute(sql)
    return cur.fetchall()

@functools.cache
def AddData(values):
    cur = con.cursor()
    cur.execute('''
    INSERT INTO `phones`(name,family,phone1,phone2,phone3,home1,home2,work_number,home_path,fax,website,email,messager,phone_msg,workpath)
    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
    ''',values)
    con.commit()
    return True

def main():
    global con
    con = db.connect('.phones.sqlite3')
    CreateTable()
    mainApp = QApplication(argv)
    mainWindow = App()
    mainApp.exec_()
    con.close()

if __name__ == "__main__" : main()
