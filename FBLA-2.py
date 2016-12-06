import sys, csv
from PyQt4 import QtGui, QtCore


class Window(QtGui.QWidget):
    def __init__(self, rows, columns):
        QtGui.QWidget.__init__(self)

        self.setWindowTitle("SpreadSheet")
        self.table = QtGui.QTableView(self)
        model =  QtGui.QStandardItemModel(rows, columns, self.table)
        self.table.setModel(model)

        # Menu Bar

        self.myQMenuBar = QtGui.QMenuBar(self)
        fileMenu = self.myQMenuBar.addMenu('File')

        extractAction = QtGui.QAction('Save', self)
        extractAction.setShortcut("Ctrl+S")
        extractAction.setStatusTip("Save File")
        extractAction.triggered.connect(self.handleSave)
        fileMenu.addAction(extractAction)

        extractAction = QtGui.QAction('Open', self)
        extractAction.setShortcut("Ctrl+O")
        extractAction.setStatusTip("Open CSV File")
        extractAction.triggered.connect(self.handleOpen)
        fileMenu.addAction(extractAction)

        extractAction = QtGui.QAction('Print', self)
        extractAction.setShortcut("Ctrl+P")
        extractAction.setStatusTip("Print Your CSV File")
        extractAction.triggered.connect(self.handlePrint)
        fileMenu.addAction(extractAction)

        extractAction = QtGui.QAction('Preview', self)
        extractAction.setShortcut("Ctrl+L")
        extractAction.setStatusTip("Preview Your CSV File")
        extractAction.triggered.connect(self.handlePreview)
        fileMenu.addAction(extractAction)         

        extractAction = QtGui.QAction('Add Row', self)
        extractAction.setShortcut("Ctrl+K")
        extractAction.setStatusTip("Add A Row")
        extractAction.triggered.connect(self.addrow)
        fileMenu.addAction(extractAction)  

        extractAction = QtGui.QAction('Remove Row', self)
        extractAction.setShortcut("Ctrl+H")
        extractAction.setStatusTip("Remove A Row")
        extractAction.triggered.connect(self.removerow)
        fileMenu.addAction(extractAction)

        extractAction = QtGui.QAction('Add Column', self)
        extractAction.setShortcut("Ctrl+J")
        extractAction.setStatusTip("Add A Column")
        extractAction.triggered.connect(self.addcolumn)
        fileMenu.addAction(extractAction)            
     
        extractAction = QtGui.QAction('Remove Column', self)
        extractAction.setShortcut("Ctrl+G")
        extractAction.setStatusTip("Remove A Column")
        extractAction.triggered.connect(self.removecolumn)
        fileMenu.addAction(extractAction)

        extractAction = QtGui.QAction('Exit', self)
        extractAction.setShortcut("Ctrl+Q")
        extractAction.setStatusTip("Leave The App")
        extractAction.triggered.connect(self.close_application)
        fileMenu.addAction(extractAction)
       
        # Table Layout

        layout = QtGui.QGridLayout(self)
        layout.addWidget(self.table, 35,35,35,35)
        self.myQMenuBar.move(0,0)

    # Adds/Removes Columns and Rows

    def removerow(self):
        selected = self.table.currentRow()
        self.table.removeRow(selected)

    def removecolumn(self):
        selected = self.table.currentColumn()
        self.table.removeColumn(selected)

    def addrow(self):
        rowPosition = self.table.rowCount()
        self.table.insertRow(rowPosition)

    def addcolumn(self):    
        columnPosition = self.table.columnCount()
        self.table.insertColumn(columnPosition)

    # Close Application

    def close_application(self):
        choice = QtGui.QMessageBox.question(self, "Warning", "Are you sure you want to exit?", QtGui.QMessageBox.Yes | QtGui.QMessageBox.No)
        
        if choice == QtGui.QMessageBox.Yes:
            sys.exit()
        else:
            pass

    # Saves File

    def handleSave(self):
        path = QtGui.QFileDialog.getSaveFileName(self, 'Save File', '', 'CSV(*.csv)')
        if not path.isEmpty():
            with open(unicode(path), 'wb') as stream:
                writer = csv.writer(stream)
                for row in range(self.table.rowCount()):
                    rowdata = []
                    for column in range(self.table.columnCount()):
                        item = self.table.item(row, column)
                        if item is not None:
                            rowdata.append(unicode(item.text()).encode('utf8'))
                        else:
                            rowdata.append('')
                    writer.writerow(rowdata)

    # Opens File

    def handleOpen(self):
        path = QtGui.QFileDialog.getOpenFileName(self, 'Open File', '', 'CSV(*.csv)')
        if not path.isEmpty():
            with open(unicode(path), 'rb') as stream:
                self.table.setRowCount(0)
                self.table.setColumnCount(0)
                for rowdata in csv.reader(stream):
                    row = self.table.rowCount()
                    self.table.insertRow(row)
                    self.table.setColumnCount(len(rowdata))
                    for column, data in enumerate(rowdata):
                        item = QtGui.QTableWidgetItem(data.decode('utf8'))
                        self.table.setItem(row, column, item)

    # Prints File

    def handlePrint(self):
        dialog = QtGui.QPrintDialog()
        if dialog.exec_() == QtGui.QDialog.Accepted:
            self.handlePaintRequest(dialog.printer())

    # Previews Table

    def handlePreview(self):
        dialog = QtGui.QPrintPreviewDialog()
        dialog.paintRequested.connect(self.handlePaintRequest)
        dialog.exec_()

    # Paints Table

    def handlePaintRequest(self, printer):
        document = QtGui.QTextDocument()
        cursor = QtGui.QTextCursor(document)
        model = self.table.model()
        table = cursor.insertTable(
            model.rowCount(), model.columnCount())
        for row in range(table.rows()):
            for column in range(table.columns()):
                cursor.insertText(self.table(row, column).text())
                cursor.movePosition(QtGui.QTextCursor.NextCell)
        document.print_(printer)

if __name__ == '__main__':

    app = QtGui.QApplication(sys.argv)
    window = Window(5, 5)
    window.resize(300, 400)
    window.show()
    sys.exit(app.exec_())
