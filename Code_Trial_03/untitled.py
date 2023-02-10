import json
import time
from io import open as opn
from label_printer import *
from Read_Excel_Stuff import *
from Secondary_Functions import *
from Write_Excel_Stuff import *
import sys
import numpy as np
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5 import QtWidgets, QtCore
global filteredRows
global selectedRow
global totalBoxes
global totalWeight
global remainingBoxQtyInEntry
global remainingWeightInEntry
global boxQty
global responseArray
global scannedLabelArray
global Excel_Read_Path
global Excel_02_Write_Path
global Excel_02_Name
global Backup_Path
global EPF
global entryFoundAt

def load_information():
    global Excel_Read_Path
    global Excel_02_Write_Path
    global Excel_02_Name
    global Backup_Path
    with opn('Required_Files/information.json', 'r') as infoRaw:
        info = json.load(infoRaw)
        Excel_Read_Path = str(info["General"]["Excel_01_Read_Path"])
        Excel_02_Write_Path = str(info["General"]["Excel_02_Write_Path"])
        Excel_02_Name = str(info["General"]["Excel_02_Name"])
        Backup_Path = str(info["General"]["Backup_Path"])


def primaryFunctions():
    Create_Excel_Document(Excel_02_Write_Path, Excel_02_Name)
    Generate_Backup_Data(Backup_Path)

class ExcelToTable(QWidget):
    def __init__(self):
        super(ExcelToTable, self).__init__()
        self.setWindowTitle("Excel Records")
        self.desktop = QApplication.desktop()
        self.screenRect = self.desktop.screenGeometry()
        self.height = int(self.screenRect.height()*0.8)
        self.width = int(self.screenRect.width()*0.8)
        self.setGeometry(int(self.width*0.1), int(self.height*0.1), self.width, self.height)
        self.initUI()
        self.loadData()

    def initUI(self):
        self.ExcelToTableLabel = QLabel(self)
        X_pos, Y_pos, L, H, Text_size = self.translate_pos(0, 0, 1, 0.1, 0.012)
        self.ExcelToTableLabel.setFixedSize(L, H)
        self.ExcelToTableLabel.move(X_pos, Y_pos)
        self.ExcelToTableLabel.setAlignment(QtCore.Qt.AlignCenter)
        self.ExcelToTableLabel.setFont(QFont('Times', Text_size))
        self.ExcelToTableLabel.setStyleSheet('color:Black')
        self.ExcelToTableLabel.setWordWrap(True)
        self.ExcelToTableLabel.setText("Excel Records")

        self.table = QtWidgets.QTableWidget(self)
        X_pos, Y_pos, L, H, Text_size = self.translate_pos(0.025, 0.08, 0.95, 0.8, 0.0185)
        self.table.setGeometry(QtCore.QRect(X_pos, Y_pos, L, H))

    def loadData(self):
        success, Rows = Open_Excel_All_Records(Excel_Read_Path)
        if success:
            numRows = len(Rows.index)
            self.table.setRowCount(0)
            self.table.clear()
            self.table.setRowCount(numRows)
            self.table.setColumnCount(12)
            self.table.setHorizontalHeaderLabels(
                ['Enter Date(MM/DD)','Enter By (Name)','Goods IN-housed Date','PO #', 'INVOICE#', 'LOT NUMBER', 'Material #', 'Twist', 'Qty Box', 'Qty Weight', 'Supplier',
                 'Customer'])
            currentRow = 0
            while (currentRow < numRows):
                self.table.setItem(currentRow, 0, QtWidgets.QTableWidgetItem(str(Rows.iloc[currentRow].at['Enter Date(MM/DD)'])))
                self.table.setItem(currentRow, 1, QtWidgets.QTableWidgetItem(str(Rows.iloc[currentRow].at['Enter By (Name)'])))
                self.table.setItem(currentRow, 2, QtWidgets.QTableWidgetItem(str(Rows.iloc[currentRow].at['Goods IN-housed Date'])))
                self.table.setItem(currentRow, 3, QtWidgets.QTableWidgetItem(str(Rows.iloc[currentRow].at['PO #'])))
                self.table.setItem(currentRow, 4, QtWidgets.QTableWidgetItem(str(Rows.iloc[currentRow].at['INVOICE#'])))
                self.table.setItem(currentRow, 5, QtWidgets.QTableWidgetItem(str(Rows.iloc[currentRow].at['LOT NUMBER'])))
                self.table.setItem(currentRow, 6, QtWidgets.QTableWidgetItem(str(Rows.iloc[currentRow].at['Material #'])))
                self.table.setItem(currentRow, 7, QtWidgets.QTableWidgetItem(str(Rows.iloc[currentRow].at['Twist'])))
                self.table.setItem(currentRow, 8, QtWidgets.QTableWidgetItem(str(Rows.iloc[currentRow].at['Qty Box'])))
                self.table.setItem(currentRow, 9, QtWidgets.QTableWidgetItem(str(Rows.iloc[currentRow].at['Qty Weight'])))
                self.table.setItem(currentRow, 10, QtWidgets.QTableWidgetItem(str(Rows.iloc[currentRow].at['SUPPLIER'])))
                self.table.setItem(currentRow, 11, QtWidgets.QTableWidgetItem(str(Rows.iloc[currentRow].at['Customer'])))
                currentRow = currentRow + 1

        else:
            QMessageBox.information(self, "Error", Rows, QMessageBox.Ok)


    def translate_pos(self, ratioX, ratioY, ratioLength, ratioHeight, ratioTextSize):
        requiredX = np.int0(ratioX*self.width)
        requiredY = np.int0(ratioY*self.height)
        requiredLength = np.int0(ratioLength*self.width)
        requiredHeight = np.int0(ratioHeight*self.height)
        requiredTextSize = np.int0(ratioTextSize*self.width)
        return requiredX, requiredY, requiredLength, requiredHeight, requiredTextSize


class ScanLabelBarcode(QWidget):
    currentScannedBarcode = None
    currentBoxQty = None
    currentWeight = None
    def __init__(self):
        super(ScanLabelBarcode, self).__init__()
        self.setWindowTitle("Scan label barcode")
        self.desktop = QApplication.desktop()
        self.screenRect = self.desktop.screenGeometry()
        self.height = int(self.screenRect.height()/1.5)
        self.width = int(self.screenRect.width()/1.5)
        self.setGeometry(int(self.width/4), int(self.height/4), self.width, self.height)
        self.initUI()

    def initUI(self):
        self.imgLabel = QLabel(self)
        X_pos, Y_pos, L, H, Text_size = self.translate_pos(0.68, 0, 0.3, 0.075, 0.0185)
        self.pixmap_raw = QPixmap('Required_Files/Autonomation_Icon.png')
        self.pixmap = self.pixmap_raw.scaled(L,H,QtCore.Qt.KeepAspectRatio)
        self.imgLabel.setPixmap(self.pixmap)
        self.imgLabel.resize(self.pixmap.width(), self.pixmap.height())
        self.imgLabel.move(X_pos,Y_pos)


        self.scanBarcodeLabel = QLabel(self)
        X_pos, Y_pos, L, H, Text_size = self.translate_pos(0.05, 0.015, 0.9, 0.1, 0.0185)
        self.scanBarcodeLabel.setFixedSize(L, H)
        self.scanBarcodeLabel.move(X_pos, Y_pos)
        self.scanBarcodeLabel.setAlignment(QtCore.Qt.AlignCenter)
        self.scanBarcodeLabel.setFont(QFont('Times', Text_size))
        self.scanBarcodeLabel.setStyleSheet('color:Black')
        self.scanBarcodeLabel.setWordWrap(True)
        self.scanBarcodeLabel.setText("SCAN LABEL BARCODE")


        self.scanBarcodeInfoLabel_01 = QLabel(self)
        X_pos, Y_pos, L, H, Text_size = self.translate_pos(0.05, 0.14, 0.4, 0.06, 0.02)
        self.scanBarcodeInfoLabel_01.setFixedSize(L, H)
        self.scanBarcodeInfoLabel_01.move(X_pos, Y_pos)
        self.scanBarcodeInfoLabel_01.setAlignment(QtCore.Qt.AlignLeft)
        self.scanBarcodeInfoLabel_01.setFont(QFont('Times', Text_size))
        self.scanBarcodeInfoLabel_01.setStyleSheet('color:Black')
        self.scanBarcodeInfoLabel_01.setWordWrap(True)
        self.scanBarcodeInfoLabel_01.setText("SCAN BARCODE HERE:")

        self.scanBarcodeButton = QPushButton(self)
        X_pos, Y_pos, L, H, Text_size = self.translate_pos(0.82, 0.132, 0.15, 0.1, 0.02)
        self.scanBarcodeButton.setFixedSize(L, H)
        self.scanBarcodeButton.move(X_pos, Y_pos)
        self.scanBarcodeButton.setFont(QFont('Times', Text_size))
        self.scanBarcodeButton.setText("Enter")
        self.scanBarcodeButton.clicked.connect(self.getBarcodeData)
        self.scanBarcodeButton.setVisible(False)

        self.scanBarcodeTextBox = QLineEdit(self)
        X_pos, Y_pos, L, H, Text_size = self.translate_pos(0.405, 0.132, 0.55, 0.1, 0.02)
        self.scanBarcodeTextBox.setFixedSize(L,H)
        self.scanBarcodeTextBox.move(X_pos,Y_pos)
        self.scanBarcodeTextBox.setFont(QFont('Times', Text_size))
        self.scanBarcodeTextBox.setPlaceholderText('Scan barcode here')
        self.scanBarcodeTextBox.setStyleSheet('color:Blue')
        self.scanBarcodeTextBox.returnPressed.connect(self.scanBarcodeButton.click)

        self.scanBarcodeInfoLabel_02 = QLabel(self)
        X_pos, Y_pos, L, H, Text_size = self.translate_pos(0.05, 0.25, 0.9, 0.085, 0.012)
        self.scanBarcodeInfoLabel_02.setFixedSize(L, H)
        self.scanBarcodeInfoLabel_02.move(X_pos, Y_pos)
        self.scanBarcodeInfoLabel_02.setAlignment(QtCore.Qt.AlignLeft)
        self.scanBarcodeInfoLabel_02.setFont(QFont('Times', Text_size))
        self.scanBarcodeInfoLabel_02.setStyleSheet('color:Red')
        self.scanBarcodeInfoLabel_02.setWordWrap(True)
        self.scanBarcodeInfoLabel_02.setText("Current scanned barcode : None")

        self.scanBarcodeInfoLabel_03 = QLabel(self)
        X_pos, Y_pos, L, H, Text_size = self.translate_pos(0.05, 0.3, 0.9, 0.085, 0.012)
        self.scanBarcodeInfoLabel_03.setFixedSize(L, H)
        self.scanBarcodeInfoLabel_03.move(X_pos, Y_pos)
        self.scanBarcodeInfoLabel_03.setAlignment(QtCore.Qt.AlignLeft)
        self.scanBarcodeInfoLabel_03.setFont(QFont('Times', Text_size))
        self.scanBarcodeInfoLabel_03.setStyleSheet('color:Red')
        self.scanBarcodeInfoLabel_03.setWordWrap(True)
        self.scanBarcodeInfoLabel_03.setText("Current box qty in pellet  : None")

        self.scanBarcodeInfoLabel_04 = QLabel(self)
        X_pos, Y_pos, L, H, Text_size = self.translate_pos(0.05, 0.35, 0.9, 0.085, 0.012)
        self.scanBarcodeInfoLabel_04.setFixedSize(L, H)
        self.scanBarcodeInfoLabel_04.move(X_pos, Y_pos)
        self.scanBarcodeInfoLabel_04.setAlignment(QtCore.Qt.AlignLeft)
        self.scanBarcodeInfoLabel_04.setFont(QFont('Times', Text_size))
        self.scanBarcodeInfoLabel_04.setStyleSheet('color:Red')
        self.scanBarcodeInfoLabel_04.setWordWrap(True)
        self.scanBarcodeInfoLabel_04.setText("Current weight in pellet   : None")

        self.scanBarcodeInfoLabel_05 = QLabel(self)
        X_pos, Y_pos, L, H, Text_size = self.translate_pos(0.05, 0.45, 0.5, 0.06, 0.016)
        self.scanBarcodeInfoLabel_05.setFixedSize(L, H)
        self.scanBarcodeInfoLabel_05.move(X_pos, Y_pos)
        self.scanBarcodeInfoLabel_05.setAlignment(QtCore.Qt.AlignLeft)
        self.scanBarcodeInfoLabel_05.setFont(QFont('Times', Text_size))
        self.scanBarcodeInfoLabel_05.setStyleSheet('color:Black')
        self.scanBarcodeInfoLabel_05.setWordWrap(True)
        self.scanBarcodeInfoLabel_05.setText("Number of boxes removed from the pellet:")

        self.removedNumBoxesTextBox = QLineEdit(self)
        X_pos, Y_pos, L, H, Text_size = self.translate_pos(0.55, 0.44, 0.13, 0.08, 0.016)
        self.removedNumBoxesTextBox.setFixedSize(L,H)
        self.removedNumBoxesTextBox.move(X_pos,Y_pos)
        self.removedNumBoxesTextBox.setFont(QFont('Times', Text_size))
        self.removedNumBoxesTextBox.setPlaceholderText('Number')
        self.removedNumBoxesTextBox.setStyleSheet('color:Blue')
        self.onlyInt = QIntValidator()
        self.removedNumBoxesTextBox.setValidator(self.onlyInt)
        self.removedNumBoxesTextBox.setEnabled(False)

        self.scanBarcodeInfoLabel_06 = QLabel(self)
        X_pos, Y_pos, L, H, Text_size = self.translate_pos(0.05, 0.55, 0.5, 0.06, 0.016)
        self.scanBarcodeInfoLabel_06.setFixedSize(L, H)
        self.scanBarcodeInfoLabel_06.move(X_pos, Y_pos)
        self.scanBarcodeInfoLabel_06.setAlignment(QtCore.Qt.AlignLeft)
        self.scanBarcodeInfoLabel_06.setFont(QFont('Times', Text_size))
        self.scanBarcodeInfoLabel_06.setStyleSheet('color:Black')
        self.scanBarcodeInfoLabel_06.setWordWrap(True)
        self.scanBarcodeInfoLabel_06.setText("Weight (Kg) removed from the pellet:")

        self.removedWeightTextBox = QLineEdit(self)
        X_pos, Y_pos, L, H, Text_size = self.translate_pos(0.55, 0.54, 0.13, 0.08, 0.016)
        self.removedWeightTextBox.setFixedSize(L,H)
        self.removedWeightTextBox.move(X_pos,Y_pos)
        self.removedWeightTextBox.setFont(QFont('Times', Text_size))
        self.removedWeightTextBox.setPlaceholderText('Kg')
        self.removedWeightTextBox.setStyleSheet('color:Blue')
        self.onlyFloat = QDoubleValidator()
        self.removedWeightTextBox.setValidator(self.onlyFloat)
        self.removedWeightTextBox.setEnabled(False)

        self.scanBarcodeInfoLabel_07 = QLabel(self)
        X_pos, Y_pos, L, H, Text_size = self.translate_pos(0.05, 0.65, 0.9, 0.085, 0.012)
        self.scanBarcodeInfoLabel_07.setFixedSize(L, H)
        self.scanBarcodeInfoLabel_07.move(X_pos, Y_pos)
        self.scanBarcodeInfoLabel_07.setAlignment(QtCore.Qt.AlignLeft)
        self.scanBarcodeInfoLabel_07.setFont(QFont('Times', Text_size))
        self.scanBarcodeInfoLabel_07.setStyleSheet('color:Red')
        self.scanBarcodeInfoLabel_07.setWordWrap(True)
        self.scanBarcodeInfoLabel_07.setText("Enter the number of boxes and weight removed from the pellet. Click the below button to save changes and generate the new label.")


        self.scanBarcodeprintLabelsButton = QPushButton(self)
        X_pos, Y_pos, L, H, Text_size = self.translate_pos(0.05, 0.9, 0.45, 0.06, 0.012)
        self.scanBarcodeprintLabelsButton.setFixedSize(L, H)
        self.scanBarcodeprintLabelsButton.move(X_pos, Y_pos)
        self.scanBarcodeprintLabelsButton.setFont(QFont('Times', Text_size))
        self.scanBarcodeprintLabelsButton.setText("Proceed and print labels")
        self.scanBarcodeprintLabelsButton.clicked.connect(self.scanBarcodePrintLabels)
        self.scanBarcodeprintLabelsButton.setEnabled(False)

        self.breakToPelletsBackButton = QPushButton(self)
        X_pos, Y_pos, L, H, Text_size = self.translate_pos(0.5, 0.9, 0.45, 0.06, 0.012)
        self.breakToPelletsBackButton.setFixedSize(L, H)
        self.breakToPelletsBackButton.move(X_pos, Y_pos)
        self.breakToPelletsBackButton.setFont(QFont('Times', Text_size))
        self.breakToPelletsBackButton.setText("Back")
        self.breakToPelletsBackButton.clicked.connect(self.back)

    def back(self):
        self.close()

    def scanBarcodePrintLabels(self):
        if self.removedNumBoxesTextBox.text() == "" or self.removedWeightTextBox.text() == "":
            QMessageBox.information(self, "Error", "Please fill in num. boxes to be removed and weight to be removed.", QMessageBox.Ok)
        else:
            barcodeInputDecoded = scannedLabelArray[0][4].split('_')
            barcodeWithoutLabelIssueNumber = barcodeInputDecoded[0] + '_' + barcodeInputDecoded[1] + '_' + barcodeInputDecoded[2]
            currentLabelIssueNumber = int(barcodeInputDecoded[3])
            newLabelIssueNumber = currentLabelIssueNumber + 1
            newOverallBarcode = barcodeWithoutLabelIssueNumber + "_" + str(newLabelIssueNumber)
            currentBoxQtyInPellet = int(scannedLabelArray[0][17])
            currentWeightInPellet = float(scannedLabelArray[0][23])
            boxQtyRemoved = int(self.removedNumBoxesTextBox.text())
            weightRemoved = float(self.removedWeightTextBox.text())

            remainingBoxQty = currentBoxQtyInPellet - boxQtyRemoved
            remainingWeight = currentWeightInPellet - weightRemoved
            remainingWeight = round(remainingWeight, 3)
            if (remainingBoxQty < 0 or remainingWeight < 0.00):
                QMessageBox.information(self, "Error", "Cannot have a negative box quantity or a negative weight.", QMessageBox.Ok)
            else:
                self.scanBarcodeprintLabelsButton.setText("Printing in progress...")
                message = 'Remaining box quantity in pellet: '+ str(remainingBoxQty) + '. remaining weight in pellet: '+ str(remainingWeight) + '. If this is OK, press Yes. Else press Cancel.'
                reply = QMessageBox.question(self, "Verify before proceeding", message,  QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
                if reply == QMessageBox.No:
                    self.scanBarcodeprintLabelsButton.setText("Proceed and print labels")
                    self.removedNumBoxesTextBox.clear()
                    self.removedWeightTextBox.clear()

                else:
                    successUpdateExistingRow = False
                    if(entryFoundAt == 0):
                        successUpdateExistingRow = Update_Existing_Row(scannedLabelArray[0][4], scannedLabelArray[2])
                    elif (entryFoundAt == 1):
                        successUpdateExistingRow = Update_Existing_Row_Previous(scannedLabelArray[0][4], scannedLabelArray[2],get_previous_month(Excel_02_Name))
                    else:
                        successUpdateExistingRow = Update_Existing_Row_Previous_Previous(scannedLabelArray[0][4], scannedLabelArray[2],get_previous_previous_month(Excel_02_Name))
                    if successUpdateExistingRow:
                        excelWriteSuccess = write_to_last_column_workbook(
                            TypeOfRecord="NEW",
                            DataEnteredBy=EPF,
                            OverallBarcode=newOverallBarcode,
                            PONumber=scannedLabelArray[0][5],
                            Excel01Row=scannedLabelArray[0][6],
                            PelletNumber=scannedLabelArray[0][7],
                            LabelIssueNumber=newLabelIssueNumber,
                            Supplier=scannedLabelArray[0][9],
                            MaterialDescription=scannedLabelArray[0][10],
                            MaterialCode=scannedLabelArray[0][11],
                            LotNumber=scannedLabelArray[0][12],
                            InvoiceNumber=scannedLabelArray[0][13],
                            TotalInitialBoxQty=scannedLabelArray[0][14],
                            BoxQtyInPelletBeforeScanning=currentBoxQtyInPellet,
                            BoxQuantityRemoved=boxQtyRemoved,
                            NewBoxQtyInPellet=remainingBoxQty,
                            Customer=scannedLabelArray[0][18],
                            Twist=scannedLabelArray[0][19],
                            WeightInPelletBeforeScanning=currentWeightInPellet,
                            WeightRemoved=weightRemoved,
                            NewWeightInPellet=remainingWeight,
                            TotalInitialWeight=scannedLabelArray[0][20]
                        )
                        backupWriteSuccess = Write_To_Backup_text_Doc(
                            TypeOfRecord="NEW",
                            DataEnteredBy=EPF,
                            OverallBarcode=newOverallBarcode,
                            PONumber=scannedLabelArray[0][5],
                            Excel01Row=scannedLabelArray[0][6],
                            PelletNumber=scannedLabelArray[0][7],
                            LabelIssueNumber=newLabelIssueNumber,
                            Supplier=scannedLabelArray[0][9],
                            MaterialDescription=scannedLabelArray[0][10],
                            MaterialCode=scannedLabelArray[0][11],
                            LotNumber=scannedLabelArray[0][12],
                            InvoiceNumber=scannedLabelArray[0][13],
                            TotalInitialBoxQty=scannedLabelArray[0][14],
                            BoxQtyInPelletBeforeScanning=currentBoxQtyInPellet,
                            BoxQuantityRemoved=boxQtyRemoved,
                            NewBoxQtyInPellet=remainingBoxQty,
                            Customer=scannedLabelArray[0][18],
                            Twist=scannedLabelArray[0][19],
                            WeightInPelletBeforeScanning=currentWeightInPellet,
                            WeightRemoved=weightRemoved,
                            NewWeightInPellet=remainingWeight,
                            TotalInitialWeight=scannedLabelArray[0][20]
                        )
                        if excelWriteSuccess == True and backupWriteSuccess == True:
                            Print_Labels(str(newOverallBarcode), str(scannedLabelArray[0][10]), str(scannedLabelArray[0][19]), str(scannedLabelArray[0][12]), str(scannedLabelArray[0][11]), str(remainingBoxQty), str(remainingWeight), str(scannedLabelArray[0][13]), str(scannedLabelArray[0][7]), str(scannedLabelArray[0][18]), str(scannedLabelArray[0][9]))
                            self.scanBarcodeprintLabelsButton.setText("Proceed and print labels")
                            self.removedNumBoxesTextBox.clear()
                            self.removedWeightTextBox.clear()
                            self.scanBarcodeInfoLabel_03.setText("Current box qty in pellet  : None")
                            self.scanBarcodeInfoLabel_04.setText("Current weight in pellet   : None")
                        else:
                            QMessageBox.information(self, "Error",
                                                    "Could not update Excel 02 and backup log file. Cannot proceed with logging information.",
                                                    QMessageBox.Ok)


                    else:
                        QMessageBox.information(self, "Error",
                                                "Could not update Excel 02. Cannot proceed with logging information.",QMessageBox.Ok)



    def getBarcodeData(self):
        global scannedLabelArray
        global entryFoundAt
        scannedLabelArray = []
        currentScannedBarcode = self.scanBarcodeTextBox.text()
        self.scanBarcodeInfoLabel_02.setText("Current scanned barcode : "+ currentScannedBarcode)
        self.scanBarcodeTextBox.clear()
        entryFoundAt = 0
        try_01 = Open_Excel_02(get_month(Excel_02_Name), str(currentScannedBarcode))
        if try_01[1]:
            scannedLabelArray = try_01
            entryFoundAt = 0
        else:
            try_02 = Open_Excel_02(get_previous_month(Excel_02_Name), str(currentScannedBarcode))
            if try_02[1]:
                scannedLabelArray = try_02
                entryFoundAt = 1
            else:
                try_03 = Open_Excel_02(get_previous_previous_month(Excel_02_Name), str(currentScannedBarcode))
                scannedLabelArray = try_03
                entryFoundAt = 2
        print(scannedLabelArray)
        if scannedLabelArray[1]:
            typeOfRecord = scannedLabelArray[0][2]
            if typeOfRecord == "NEW":
                self.removedNumBoxesTextBox.setEnabled(True)
                self.removedWeightTextBox.setEnabled(True)
                self.scanBarcodeprintLabelsButton.setEnabled(True)
                res1 = "Current box qty in pellet  : " + str(scannedLabelArray[0][17])
                res2 = "Current weight in pellet   : " + str(scannedLabelArray[0][23])
                self.scanBarcodeInfoLabel_03.setText(res1)
                self.scanBarcodeInfoLabel_04.setText(res2)

            else:
                QMessageBox.information(self, "Error", "This label is USED. Cannot scan again.")
                self.removedNumBoxesTextBox.setEnabled(False)
                self.removedWeightTextBox.setEnabled(False)
                self.scanBarcodeprintLabelsButton.setEnabled(False)
                self.scanBarcodeInfoLabel_03.setText("Current box qty in pellet  : None")
                self.scanBarcodeInfoLabel_04.setText("Current weight in pellet   : None")


        else:
            QMessageBox.information(self, "Error", str(scannedLabelArray[2]), QMessageBox.Ok)
            self.removedNumBoxesTextBox.setEnabled(False)
            self.removedWeightTextBox.setEnabled(False)
            self.scanBarcodeprintLabelsButton.setEnabled(False)
            self.scanBarcodeInfoLabel_03.setText("Current box qty in pellet  : None")
            self.scanBarcodeInfoLabel_04.setText("Current weight in pellet   : None")





    def translate_pos(self, ratioX, ratioY, ratioLength, ratioHeight, ratioTextSize):
        requiredX = np.int0(ratioX*self.width)
        requiredY = np.int0(ratioY*self.height)
        requiredLength = np.int0(ratioLength*self.width)
        requiredHeight = np.int0(ratioHeight*self.height)
        requiredTextSize = np.int0(ratioTextSize*self.width)
        return requiredX, requiredY, requiredLength, requiredHeight, requiredTextSize

class BreakIntoPellets(QWidget):

    def __init__(self):
        super(BreakIntoPellets, self).__init__()
        self.setWindowTitle("Break into pellets (Row in Excel_01: "+str(selectedExcelRow)+")")
        self.desktop = QApplication.desktop()
        self.screenRect = self.desktop.screenGeometry()
        self.height = int(self.screenRect.height()/1.5)
        self.width = int(self.screenRect.width()/1.5)
        self.setGeometry(int(self.width/4), int(self.height/4), self.width, self.height)
        self.initUI()

    def initUI(self):
        global filteredRows
        global selectedExcelRow
        global totalBoxes
        global totalWeight
        global remainingBoxQtyInEntry
        global remainingWeightInEntry
        global boxQty
        global responseArray

        totalBoxes = int(filteredRows.iloc[selectedRow].at['Qty Box'])
        totalWeight = float(filteredRows.iloc[selectedRow].at['Qty Weight'])
        remainingBoxQtyInEntry = totalBoxes
        remainingWeightInEntry = totalWeight
        boxQty = 0
        responseArray = []

        self.imgLabel = QLabel(self)
        X_pos, Y_pos, L, H, Text_size = self.translate_pos(0.68, 0, 0.3, 0.075, 0.0185)
        self.pixmap_raw = QPixmap('Required_Files/Autonomation_Icon.png')
        self.pixmap = self.pixmap_raw.scaled(L,H,QtCore.Qt.KeepAspectRatio)
        self.imgLabel.setPixmap(self.pixmap)
        self.imgLabel.resize(self.pixmap.width(), self.pixmap.height())
        self.imgLabel.move(X_pos,Y_pos)


        self.BreakIntoPelletsLabel = QLabel(self)
        X_pos, Y_pos, L, H, Text_size = self.translate_pos(0.05, 0.015, 0.9, 0.1, 0.0185)
        self.BreakIntoPelletsLabel.setFixedSize(L, H)
        self.BreakIntoPelletsLabel.move(X_pos, Y_pos)
        self.BreakIntoPelletsLabel.setAlignment(QtCore.Qt.AlignCenter)
        self.BreakIntoPelletsLabel.setFont(QFont('Times', Text_size))
        self.BreakIntoPelletsLabel.setStyleSheet('color:Black')
        self.BreakIntoPelletsLabel.setWordWrap(True)
        self.BreakIntoPelletsLabel.setText("BREAK INTO PELLETS")

        self.BreakIntoPelletsInfoLabel_01 = QLabel(self)
        X_pos, Y_pos, L, H, Text_size = self.translate_pos(0.05, 0.1, 0.9, 0.085, 0.012)
        self.BreakIntoPelletsInfoLabel_01.setFixedSize(L, H)
        self.BreakIntoPelletsInfoLabel_01.move(X_pos, Y_pos)
        self.BreakIntoPelletsInfoLabel_01.setAlignment(QtCore.Qt.AlignLeft)
        self.BreakIntoPelletsInfoLabel_01.setFont(QFont('Times', Text_size))
        self.BreakIntoPelletsInfoLabel_01.setStyleSheet('color:Red')
        self.BreakIntoPelletsInfoLabel_01.setWordWrap(True)
        self.BreakIntoPelletsInfoLabel_01.setText("Remaining box qty in the entry: " + str(remainingBoxQtyInEntry))

        self.BreakIntoPelletsInfoLabel_02 = QLabel(self)
        X_pos, Y_pos, L, H, Text_size = self.translate_pos(0.05, 0.15, 0.9, 0.085, 0.012)
        self.BreakIntoPelletsInfoLabel_02.setFixedSize(L, H)
        self.BreakIntoPelletsInfoLabel_02.move(X_pos, Y_pos)
        self.BreakIntoPelletsInfoLabel_02.setAlignment(QtCore.Qt.AlignLeft)
        self.BreakIntoPelletsInfoLabel_02.setFont(QFont('Times', Text_size))
        self.BreakIntoPelletsInfoLabel_02.setStyleSheet('color:Red')
        self.BreakIntoPelletsInfoLabel_02.setWordWrap(True)
        self.BreakIntoPelletsInfoLabel_02.setText("Remaining weight in the entry: " + str(remainingWeightInEntry)+ " Kg")

        self.BreakIntoPelletsInfoLabel_03 = QLabel(self)
        X_pos, Y_pos, L, H, Text_size = self.translate_pos(0.05, 0.21, 0.4, 0.06, 0.012)
        self.BreakIntoPelletsInfoLabel_03.setFixedSize(L, H)
        self.BreakIntoPelletsInfoLabel_03.move(X_pos, Y_pos)
        self.BreakIntoPelletsInfoLabel_03.setAlignment(QtCore.Qt.AlignLeft)
        self.BreakIntoPelletsInfoLabel_03.setFont(QFont('Times', Text_size))
        self.BreakIntoPelletsInfoLabel_03.setStyleSheet('color:Black')
        self.BreakIntoPelletsInfoLabel_03.setWordWrap(True)
        self.BreakIntoPelletsInfoLabel_03.setText("Number of boxes allocated to the pellet:")

        self.numBoxesTextBox = QLineEdit(self)
        X_pos, Y_pos, L, H, Text_size = self.translate_pos(0.405, 0.205, 0.17, 0.06, 0.012)
        self.numBoxesTextBox.setFixedSize(L,H)
        self.numBoxesTextBox.move(X_pos,Y_pos)
        self.numBoxesTextBox.setFont(QFont('Times', Text_size))
        self.numBoxesTextBox.setPlaceholderText('Enter num. boxes')
        self.numBoxesTextBox.setStyleSheet('color:Blue')
        self.onlyInt = QIntValidator()
        self.numBoxesTextBox.setValidator(self.onlyInt)

        self.BreakIntoPelletsInfoLabel_04 = QLabel(self)
        X_pos, Y_pos, L, H, Text_size = self.translate_pos(0.05, 0.28, 0.4, 0.06, 0.012)
        self.BreakIntoPelletsInfoLabel_04.setFixedSize(L, H)
        self.BreakIntoPelletsInfoLabel_04.move(X_pos, Y_pos)
        self.BreakIntoPelletsInfoLabel_04.setAlignment(QtCore.Qt.AlignLeft)
        self.BreakIntoPelletsInfoLabel_04.setFont(QFont('Times', Text_size))
        self.BreakIntoPelletsInfoLabel_04.setStyleSheet('color:Black')
        self.BreakIntoPelletsInfoLabel_04.setWordWrap(True)
        self.BreakIntoPelletsInfoLabel_04.setText("Total Kg in the pellet:")

        self.weightTextBox = QLineEdit(self)
        X_pos, Y_pos, L, H, Text_size = self.translate_pos(0.405, 0.275, 0.17, 0.06, 0.012)
        self.weightTextBox.setFixedSize(L,H)
        self.weightTextBox.move(X_pos,Y_pos)
        self.weightTextBox.setFont(QFont('Times', Text_size))
        self.weightTextBox.setPlaceholderText('Enter weight (Kg)')
        self.weightTextBox.setStyleSheet('color:Blue')
        self.onlyFloat = QDoubleValidator()
        self.weightTextBox.setValidator(self.onlyFloat)

        self.addNewPelletButton = QPushButton(self)
        X_pos, Y_pos, L, H, Text_size = self.translate_pos(0.05, 0.35, 0.45, 0.06, 0.012)
        self.addNewPelletButton.setFixedSize(L, H)
        self.addNewPelletButton.move(X_pos, Y_pos)
        self.addNewPelletButton.setFont(QFont('Times', Text_size))
        self.addNewPelletButton.setText("Add New Pellet")
        self.addNewPelletButton.clicked.connect(self.addNewPellet)

        self.removeSelectedPelletButton = QPushButton(self)
        X_pos, Y_pos, L, H, Text_size = self.translate_pos(0.5, 0.35, 0.45, 0.06, 0.012)
        self.removeSelectedPelletButton.setFixedSize(L, H)
        self.removeSelectedPelletButton.move(X_pos, Y_pos)
        self.removeSelectedPelletButton.setFont(QFont('Times', Text_size))
        self.removeSelectedPelletButton.setText("Remove Selected Pellet")
        self.removeSelectedPelletButton.clicked.connect(self.removeSelectedPellet)

        self.table = QtWidgets.QTableWidget(self)
        X_pos, Y_pos, L, H, Text_size = self.translate_pos(0.05, 0.43, 0.9, 0.35, 0.0185)
        self.table.setGeometry(QtCore.QRect(X_pos, Y_pos, L, H))

        self.BreakIntoPelletsInfoLabel_05 = QLabel(self)
        X_pos, Y_pos, L, H, Text_size = self.translate_pos(0.05, 0.8, 0.9, 0.085, 0.012)
        self.BreakIntoPelletsInfoLabel_05.setFixedSize(L, H)
        self.BreakIntoPelletsInfoLabel_05.move(X_pos, Y_pos)
        self.BreakIntoPelletsInfoLabel_05.setAlignment(QtCore.Qt.AlignLeft)
        self.BreakIntoPelletsInfoLabel_05.setFont(QFont('Times', Text_size))
        self.BreakIntoPelletsInfoLabel_05.setStyleSheet('color:Red')
        self.BreakIntoPelletsInfoLabel_05.setWordWrap(True)
        self.BreakIntoPelletsInfoLabel_05.setText("Above table shows information on the number of pellets created and how many boxes in each pellet. If all is ok, click the below button")

        self.printLabelsButton = QPushButton(self)
        X_pos, Y_pos, L, H, Text_size = self.translate_pos(0.05, 0.9, 0.45, 0.06, 0.012)
        self.printLabelsButton.setFixedSize(L, H)
        self.printLabelsButton.move(X_pos, Y_pos)
        self.printLabelsButton.setFont(QFont('Times', Text_size))
        self.printLabelsButton.setText("Proceed and print labels")
        self.printLabelsButton.clicked.connect(self.priorToPrintLabels)

        self.breakToPelletsBackButton = QPushButton(self)
        X_pos, Y_pos, L, H, Text_size = self.translate_pos(0.5, 0.9, 0.45, 0.06, 0.012)
        self.breakToPelletsBackButton.setFixedSize(L, H)
        self.breakToPelletsBackButton.move(X_pos, Y_pos)
        self.breakToPelletsBackButton.setFont(QFont('Times', Text_size))
        self.breakToPelletsBackButton.setText("Back")
        self.breakToPelletsBackButton.clicked.connect(self.back)


    def addNewPellet(self):
        global filteredRows
        global selectedExcelRow
        global totalBoxes
        global totalWeight
        global remainingBoxQtyInEntry
        global remainingWeightInEntry
        global boxQty
        global responseArray
        if self.numBoxesTextBox.text() == "" or self.weightTextBox.text() == "":
            QMessageBox.information(self, "Error", "Please fill in num. boxes and weight allocated.", QMessageBox.Ok)
        else:
            pelletBoxQty = int(self.numBoxesTextBox.text())
            pelletWeight = float(self.weightTextBox.text())
            remainingBoxQtyInEntry = remainingBoxQtyInEntry - pelletBoxQty
            remainingWeightInEntry = remainingWeightInEntry - pelletWeight
            if remainingBoxQtyInEntry <0 or remainingWeightInEntry <0:
                remainingBoxQtyInEntry = remainingBoxQtyInEntry + pelletBoxQty
                remainingWeightInEntry = remainingWeightInEntry + pelletWeight
                QMessageBox.information(self, "Error", "Cannot have negative values for box qty or weight.",QMessageBox.Ok)
            else:
                pelletWeight = round(pelletWeight,3)
                remainingWeightInEntry = round(remainingWeightInEntry,3)
                semiElement = [
                    (str(filteredRows.iloc[selectedRow].at['PO #'])),
                    filteredRows.iloc[selectedRow].at['INVOICE#'],
                    filteredRows.iloc[selectedRow].at['Material #'],
                    filteredRows.iloc[selectedRow].at['YARN ARTICLE'],
                    filteredRows.iloc[selectedRow].at['SUPPLIER'],
                    filteredRows.iloc[selectedRow].at['LOT NUMBER'],
                    int(filteredRows.iloc[selectedRow].at['Qty Box']),
                    filteredRows.iloc[selectedRow].at['PO #'],
                    filteredRows.index[selectedRow],
                    filteredRows.iloc[selectedRow].at['Customer'],
                    float(filteredRows.iloc[selectedRow].at['Qty Weight']),
                    filteredRows.iloc[selectedRow].at['Twist'],
                    pelletBoxQty,
                    pelletWeight
                ]
                responseArray.append(semiElement)
                self.fillDetailsToTable()
            self.numBoxesTextBox.clear()
            self.weightTextBox.clear()


    def fillDetailsToTable(self):

        numRows = len(responseArray)
        self.table.setRowCount(0)
        self.table.clear()
        self.table.setRowCount(numRows)
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(['BC ref', 'Pellet Box Qty', 'Pellet Weight'])
        currentRow = 0
        for response in responseArray:
            self.table.setItem(currentRow, 0, QtWidgets.QTableWidgetItem(str(response[0])))
            self.table.setItem(currentRow, 1, QtWidgets.QTableWidgetItem(str(response[12])))
            self.table.setItem(currentRow, 2, QtWidgets.QTableWidgetItem(str(response[13])))
            currentRow = currentRow+1

        self.BreakIntoPelletsInfoLabel_01.setText("Remaining box qty in the entry: " + str(remainingBoxQtyInEntry))
        self.BreakIntoPelletsInfoLabel_02.setText("Remaining weight in the entry: " + str(remainingWeightInEntry) + " Kg")


    def removeSelectedPellet(self):
        global remainingBoxQtyInEntry
        global remainingWeightInEntry
        if self.table.currentRow() != -1:
            remainingBoxQtyInEntry = remainingBoxQtyInEntry + responseArray[self.table.currentRow()][12]
            remainingWeightInEntry = remainingWeightInEntry + responseArray[self.table.currentRow()][13]
            responseArray.pop(self.table.currentRow())
            self.fillDetailsToTable()

        else:
            pass

    def priorToPrintLabels(self):
        self.printLabelsButton.setText("Printing in progress...")
        reply = QMessageBox.question(self, "Verify before proceeding", "Confirm label printing?", QMessageBox.Yes | QMessageBox.No,QMessageBox.Yes)
        if reply == QMessageBox.No:
            self.printLabelsButton.setText("Proceed and print labels")
        else:
            self.printLabels()

    def printLabels(self):
        global responseArray
        if len(responseArray) ==0:
            QMessageBox.information(self, "Error", "No pellets created!", QMessageBox.Ok)
            self.printLabelsButton.setText("Proceed and print labels")
        else:
            printout_success = False
            for index, response in enumerate(responseArray):
                response[0] = response[0] + "_"+str(response[8])+"_" +str(index+1)+"_1"
                pelletNum = str(index + 1)
                excelWriteSuccess = write_to_last_column_workbook(
                    TypeOfRecord="NEW",
                    DataEnteredBy=str(EPF),
                    OverallBarcode= response[0],
                    PONumber=response[7],
                    Excel01Row=response[8],
                    PelletNumber=pelletNum,
                    LabelIssueNumber='1',
                    Supplier=response[4],
                    MaterialDescription=response[3],
                    MaterialCode=str(response[2]),
                    LotNumber=response[5],
                    InvoiceNumber=response[1],
                    TotalInitialBoxQty=response[6],
                    BoxQtyInPelletBeforeScanning=response[12],
                    BoxQuantityRemoved=0,
                    NewBoxQtyInPellet=response[12],
                    Customer=response[9],
                    Twist=response[11],
                    WeightInPelletBeforeScanning=response[13],
                    WeightRemoved=0,
                    NewWeightInPellet=response[13],
                    TotalInitialWeight=response[10]
                )
                backupWriteSuccess = Write_To_Backup_text_Doc(
                    TypeOfRecord="NEW",
                    DataEnteredBy=str(EPF),
                    OverallBarcode=response[0],
                    PONumber=response[7],
                    Excel01Row=response[8],
                    PelletNumber=pelletNum,
                    LabelIssueNumber='1',
                    Supplier=response[4],
                    MaterialDescription=response[3],
                    MaterialCode=str(response[2]),
                    LotNumber=response[5],
                    InvoiceNumber=response[1],
                    TotalInitialBoxQty=response[6],
                    BoxQtyInPelletBeforeScanning=response[12],
                    BoxQuantityRemoved=0,
                    NewBoxQtyInPellet=response[12],
                    Customer=response[9],
                    Twist=response[11],
                    WeightInPelletBeforeScanning=response[13],
                    WeightRemoved=0,
                    NewWeightInPellet=response[13],
                    TotalInitialWeight=response[10]
                )
                if excelWriteSuccess == True and backupWriteSuccess == True:
                    printout_success = True
                    Print_Labels(str(response[0]), str(response[3]), str(response[11]), str(response[5]), str(response[2]), str(response[12]), str(response[13]),str(response[1]), str(pelletNum), str(response[9]), str(response[4]))
                else:
                    QMessageBox.information(self, "Error", "Could not record information in the Excel_02 or backup doc. Please try again.", QMessageBox.Ok)
                    printout_success = False
            if printout_success:
                QMessageBox.information(self, "Success","Printout generation completed.",QMessageBox.Ok)
                self.close()

    def back(self):
        print('button seleced to go back')
        self.close()


    def translate_pos(self, ratioX, ratioY, ratioLength, ratioHeight, ratioTextSize):
        requiredX = np.int0(ratioX*self.width)
        requiredY = np.int0(ratioY*self.height)
        requiredLength = np.int0(ratioLength*self.width)
        requiredHeight = np.int0(ratioHeight*self.height)
        requiredTextSize = np.int0(ratioTextSize*self.width)
        return requiredX, requiredY, requiredLength, requiredHeight, requiredTextSize


class CreateANewEntryWindow(QWidget):

    def __init__(self):
        global filteredRows
        global responseArray
        super(CreateANewEntryWindow, self).__init__()
        self.setWindowTitle("Crete a new entry")
        self.desktop = QApplication.desktop()
        self.screenRect = self.desktop.screenGeometry()
        self.height = int(self.screenRect.height()/1.5)
        self.width = int(self.screenRect.width()/1.5)
        self.setGeometry(int(self.width/4), int(self.height/4), self.width, self.height)
        self.initUI()


    def initUI(self):

        self.imgLabel = QLabel(self)
        X_pos, Y_pos, L, H, Text_size = self.translate_pos(0.68, 0, 0.3, 0.075, 0.0185)
        self.pixmap_raw = QPixmap('Required_Files/Autonomation_Icon.png')
        self.pixmap = self.pixmap_raw.scaled(L,H,QtCore.Qt.KeepAspectRatio)
        self.imgLabel.setPixmap(self.pixmap)
        self.imgLabel.resize(self.pixmap.width(), self.pixmap.height())
        self.imgLabel.move(X_pos,Y_pos)

        self.createNewEntryLabel = QLabel(self)
        X_pos, Y_pos, L, H, Text_size = self.translate_pos(0.05, 0.015, 0.9, 0.1, 0.0185)
        self.createNewEntryLabel.setFixedSize(L, H)
        self.createNewEntryLabel.move(X_pos, Y_pos)
        self.createNewEntryLabel.setAlignment(QtCore.Qt.AlignCenter)
        self.createNewEntryLabel.setFont(QFont('Times', Text_size))
        self.createNewEntryLabel.setStyleSheet('color:Black')
        self.createNewEntryLabel.setWordWrap(True)
        self.createNewEntryLabel.setText("CREATE A NEW ENTRY")

        self.createNewEntryInfoLabel_01 = QLabel(self)
        X_pos, Y_pos, L, H, Text_size = self.translate_pos(0.05, 0.1, 0.9, 0.085, 0.012)
        self.createNewEntryInfoLabel_01.setFixedSize(L, H)
        self.createNewEntryInfoLabel_01.move(X_pos, Y_pos)
        self.createNewEntryInfoLabel_01.setAlignment(QtCore.Qt.AlignLeft)
        self.createNewEntryInfoLabel_01.setFont(QFont('Times', Text_size))
        self.createNewEntryInfoLabel_01.setStyleSheet('color:Red')
        self.createNewEntryInfoLabel_01.setWordWrap(True)
        self.createNewEntryInfoLabel_01.setText("* Enter the material code or Invoice number or any preferred filter text to the textbox below. The table will display the filtered rows for you to choose from.")

        self.createNewEntryInfoLabel_02 = QLabel(self)
        X_pos, Y_pos, L, H, Text_size = self.translate_pos(0.05, 0.205, 0.3, 0.06, 0.012)
        self.createNewEntryInfoLabel_02.setFixedSize(L, H)
        self.createNewEntryInfoLabel_02.move(X_pos, Y_pos)
        self.createNewEntryInfoLabel_02.setAlignment(QtCore.Qt.AlignLeft)
        self.createNewEntryInfoLabel_02.setFont(QFont('Times', Text_size))
        self.createNewEntryInfoLabel_02.setStyleSheet('color:Black')
        self.createNewEntryInfoLabel_02.setWordWrap(True)
        self.createNewEntryInfoLabel_02.setText("Search for an item:")

        self.createNewEntryTextBox = QLineEdit(self)
        X_pos, Y_pos, L, H, Text_size = self.translate_pos(0.4, 0.205, 0.3, 0.06, 0.012)
        self.createNewEntryTextBox.setFixedSize(L,H)
        self.createNewEntryTextBox.move(X_pos,Y_pos)
        self.createNewEntryTextBox.setFont(QFont('Times', Text_size))
        self.createNewEntryTextBox.setPlaceholderText('Enter Query')
        self.createNewEntryTextBox.setStyleSheet('color:Blue')

        self.createNewEntrySearchButton = QPushButton(self)
        X_pos, Y_pos, L, H, Text_size = self.translate_pos(0.7, 0.205, 0.1, 0.06, 0.012)
        self.createNewEntrySearchButton.setFixedSize(L, H)
        self.createNewEntrySearchButton.move(X_pos, Y_pos)
        self.createNewEntrySearchButton.setFont(QFont('Times', Text_size))
        self.createNewEntrySearchButton.setText("Search")
        self.createNewEntrySearchButton.clicked.connect(self.searchQuery)

        self.createNewEntryClearSearchButton = QPushButton(self)
        X_pos, Y_pos, L, H, Text_size = self.translate_pos(0.8, 0.205, 0.1, 0.06, 0.012)
        self.createNewEntryClearSearchButton.setFixedSize(L, H)
        self.createNewEntryClearSearchButton.move(X_pos, Y_pos)
        self.createNewEntryClearSearchButton.setFont(QFont('Times', Text_size))
        self.createNewEntryClearSearchButton.setText("Clear")
        self.createNewEntryClearSearchButton.clicked.connect(self.clearQuery)

        self.table = QtWidgets.QTableWidget(self)
        X_pos, Y_pos, L, H, Text_size = self.translate_pos(0.05, 0.27, 0.9, 0.52, 0.0185)
        self.table.setGeometry(QtCore.QRect(X_pos, Y_pos, L, H))

        self.createNewEntryInfoLabel_03 = QLabel(self)
        X_pos, Y_pos, L, H, Text_size = self.translate_pos(0.05, 0.8, 0.9, 0.06, 0.012)
        self.createNewEntryInfoLabel_03.setFixedSize(L, H)
        self.createNewEntryInfoLabel_03.move(X_pos, Y_pos)
        self.createNewEntryInfoLabel_03.setAlignment(QtCore.Qt.AlignLeft)
        self.createNewEntryInfoLabel_03.setFont(QFont('Times', Text_size))
        self.createNewEntryInfoLabel_03.setStyleSheet('color:Red')
        self.createNewEntryInfoLabel_03.setWordWrap(True)
        self.createNewEntryInfoLabel_03.setText("Highlight the required entry from above table and press 'Proceed'.")

        self.createNewEntryProceedButton = QPushButton(self)
        X_pos, Y_pos, L, H, Text_size = self.translate_pos(0.1, 0.87, 0.35, 0.1, 0.012)
        self.createNewEntryProceedButton.setFixedSize(L, H)
        self.createNewEntryProceedButton.move(X_pos, Y_pos)
        self.createNewEntryProceedButton.setFont(QFont('Times', Text_size))
        self.createNewEntryProceedButton.setText("Proceed")
        self.createNewEntryProceedButton.clicked.connect(self.proceed)

        self.createNewEntryBackButton = QPushButton(self)
        X_pos, Y_pos, L, H, Text_size = self.translate_pos(0.55, 0.87, 0.35, 0.1, 0.012)
        self.createNewEntryBackButton.setFixedSize(L, H)
        self.createNewEntryBackButton.move(X_pos, Y_pos)
        self.createNewEntryBackButton.setFont(QFont('Times', Text_size))
        self.createNewEntryBackButton.setText("Back")
        self.createNewEntryBackButton.clicked.connect(self.back)

    def translate_pos(self, ratioX, ratioY, ratioLength, ratioHeight, ratioTextSize):
        requiredX = np.int0(ratioX*self.width)
        requiredY = np.int0(ratioY*self.height)
        requiredLength = np.int0(ratioLength*self.width)
        requiredHeight = np.int0(ratioHeight*self.height)
        requiredTextSize = np.int0(ratioTextSize*self.width)
        return requiredX, requiredY, requiredLength, requiredHeight, requiredTextSize

    def searchQuery(self):
        global filteredRows
        if(self.createNewEntryTextBox.text()):
            success, filteredRows = Open_Excel_01_In_UI(Excel_Read_Path, self.createNewEntryTextBox.text())
            if success:
                self.fillDetailsToTable(filteredRows)
            else:
                QMessageBox.information(self, "Error", filteredRows, QMessageBox.Ok)

        else:
            QMessageBox.information(self, "Error", "Search Query is blank", QMessageBox.Ok)

    def fillDetailsToTable(self, df):
        numRows = len(df.index)
        self.table.setRowCount(0)
        self.table.clear()
        self.table.setRowCount(numRows)
        self.table.setColumnCount(9)
        self.table.setHorizontalHeaderLabels(['PO #','INVOICE#','LOT NUMBER','Material #', 'Twist','Qty Box','Qty Weight','Supplier','Customer'])
        currentRow = 0
        while(currentRow < numRows):
            self.table.setItem(currentRow,0,QtWidgets.QTableWidgetItem(str(df.iloc[currentRow].at['PO #'])))
            self.table.setItem(currentRow, 1, QtWidgets.QTableWidgetItem(str(df.iloc[currentRow].at['INVOICE#'])))
            self.table.setItem(currentRow, 2, QtWidgets.QTableWidgetItem(str(df.iloc[currentRow].at['LOT NUMBER'])))
            self.table.setItem(currentRow, 3, QtWidgets.QTableWidgetItem(str(df.iloc[currentRow].at['Material #'])))
            self.table.setItem(currentRow, 4, QtWidgets.QTableWidgetItem(str(df.iloc[currentRow].at['Twist'])))
            self.table.setItem(currentRow, 5, QtWidgets.QTableWidgetItem(str(df.iloc[currentRow].at['Qty Box'])))
            self.table.setItem(currentRow, 6, QtWidgets.QTableWidgetItem(str(df.iloc[currentRow].at['Qty Weight'])))
            self.table.setItem(currentRow, 7, QtWidgets.QTableWidgetItem(str(df.iloc[currentRow].at['SUPPLIER'])))
            self.table.setItem(currentRow, 8, QtWidgets.QTableWidgetItem(str(df.iloc[currentRow].at['Customer'])))
            currentRow = currentRow+1
    def clearQuery(self):
        self.createNewEntryTextBox.clear()
        self.table.setRowCount(0)
        self.table.setColumnCount(0)
        self.table.clear()

    def proceed(self):
        global filteredRows
        global selectedExcelRow
        global selectedRow
        global responseArray
        if self.table.currentRow() ==-1:
            QMessageBox.information(self, "Error", "Please select a row from the table to proceed.", QMessageBox.Ok)
            pass
        else:
            selectedExcelRow = filteredRows.index[self.table.currentRow()]
            selectedRow = self.table.currentRow()
            barcodeReference = str(filteredRows.iloc[selectedRow].at['PO #']) + "_" + str(selectedExcelRow) + '_1_1'
            recordAlreadyExist1 = Check_Excel_02_Record_Duplication(get_month(Excel_02_Name), (barcodeReference))
            recordAlreadyExist2 = Check_Excel_02_Record_Duplication(get_previous_month(Excel_02_Name), (barcodeReference))
            if recordAlreadyExist1[0]:
                message = str(recordAlreadyExist1[1])
                QMessageBox.information(self,"Error",message , QMessageBox.Ok)
            else:
                pass

            if recordAlreadyExist2[0]:
                message = str(recordAlreadyExist2[1])
                QMessageBox.information(self,"Error",message , QMessageBox.Ok)
            else:
                pass


            if recordAlreadyExist1[0] == False and recordAlreadyExist2[0] == False:
                responseArray = []
                self.w = BreakIntoPellets()
                self.w.show()
                self.close()
            else:
                pass

    def back(self):
        print('button to return is pressed.')
        self.close()


class MyWindow(QMainWindow):
    def __init__(self):
        super(MyWindow, self).__init__()
        self.desktop = QApplication.desktop()
        self.screenRect = self.desktop.screenGeometry()
        self.height = int(self.screenRect.height()/2)
        self.width = int(self.screenRect.width()/2)
        self.setGeometry(int(self.width/2), int(self.height/2), self.width, self.height)
        self.setWindowTitle("Oviklo Yarn Stores Inventory Management System")
        self.initUI()


    def initUI(self):
        self.imgLabel = QLabel(self)
        X_pos, Y_pos, L, H, Text_size = self.translate_pos(0.68, 0, 0.3, 0.075, 0.0185)
        self.pixmap_raw = QPixmap('Required_Files/Autonomation_Icon.png')
        self.pixmap = self.pixmap_raw.scaled(L,H,QtCore.Qt.KeepAspectRatio)
        self.imgLabel.setPixmap(self.pixmap)
        self.imgLabel.resize(self.pixmap.width(), self.pixmap.height())
        self.imgLabel.move(X_pos,Y_pos)

        self.mainWindowLabel = QLabel(self)
        X_pos, Y_pos, L, H, Text_size = self.translate_pos(0.1, 0.05, 0.8, 0.075, 0.0185)
        self.mainWindowLabel.setFixedSize(L, H)
        self.mainWindowLabel.move(X_pos, Y_pos)
        self.mainWindowLabel.setAlignment(QtCore.Qt.AlignCenter)
        self.mainWindowLabel.setFont(QFont('Times', Text_size))
        self.mainWindowLabel.setStyleSheet('color:Black')
        self.mainWindowLabel.setText("OVIKLO INVENTORY MANAGEMENT SYSTEM")

        self.mainWindowLabel_02 = QLabel(self)
        X_pos, Y_pos, L, H, Text_size = self.translate_pos(0.3, 0.21, 0.3, 0.06, 0.012)
        self.mainWindowLabel_02.setFixedSize(L, H)
        self.mainWindowLabel_02.move(X_pos, Y_pos)
        self.mainWindowLabel_02.setAlignment(QtCore.Qt.AlignLeft)
        self.mainWindowLabel_02.setFont(QFont('Times', Text_size))
        self.mainWindowLabel_02.setStyleSheet('color:Black')
        self.mainWindowLabel_02.setWordWrap(True)
        self.mainWindowLabel_02.setText("Enter EPF:")

        self.mainWindowTextBox = QLineEdit(self)
        X_pos, Y_pos, L, H, Text_size = self.translate_pos(0.4, 0.205, 0.3, 0.06, 0.012)
        self.mainWindowTextBox.setFixedSize(L,H)
        self.mainWindowTextBox.move(X_pos,Y_pos)
        self.mainWindowTextBox.setFont(QFont('Times', Text_size))
        self.mainWindowTextBox.setPlaceholderText('EPF')
        self.mainWindowTextBox.setStyleSheet('color:Blue')
        self.mainWindowTextBox.textChanged.connect(self.EPFEntered)

        self.mainWindowInfoLabel = QLabel(self)
        X_pos, Y_pos, L, H, Text_size = self.translate_pos(0.01, 0.95, 0.8, 0.075, 0.012)
        self.mainWindowInfoLabel.setFixedSize(L, H)
        self.mainWindowInfoLabel.move(X_pos, Y_pos)
        self.mainWindowInfoLabel.setAlignment(QtCore.Qt.AlignLeft)
        self.mainWindowInfoLabel.setFont(QFont('Times', Text_size))
        self.mainWindowInfoLabel.setStyleSheet('color:Red')
        self.mainWindowInfoLabel.setText("* Note: Please Close the Excel_01 and Excel_02 documents before using the software")

        self.mainWindowBtn_createNewEntry = QtWidgets.QPushButton(self)
        X_pos, Y_pos, L, H, Text_size = self.translate_pos(0.1, 0.60, 0.35, 0.1, 0.015)
        self.mainWindowBtn_createNewEntry.setFixedSize(L, H)
        self.mainWindowBtn_createNewEntry.move(X_pos, Y_pos)
        self.mainWindowBtn_createNewEntry.setFont(QFont('Times', Text_size))
        self.mainWindowBtn_createNewEntry.setText("CREATE A NEW ENTRY")
        self.mainWindowBtn_createNewEntry.clicked.connect(self.createNewEntry)
        self.mainWindowBtn_createNewEntry.setEnabled(False)

        self.mainWindowBtn_scanALabel = QtWidgets.QPushButton(self)
        X_pos, Y_pos, L, H, Text_size = self.translate_pos(0.5, 0.60, 0.35, 0.1, 0.015)
        self.mainWindowBtn_scanALabel.setFixedSize(L, H)
        self.mainWindowBtn_scanALabel.move(X_pos, Y_pos)
        self.mainWindowBtn_scanALabel.setFont(QFont('Times', Text_size))
        self.mainWindowBtn_scanALabel.setText("SCAN A LABEL")
        self.mainWindowBtn_scanALabel.clicked.connect(self.scanALabel)
        self.mainWindowBtn_scanALabel.setEnabled(False)

        self.mainWindowBtn_excelToTable = QtWidgets.QPushButton(self)
        X_pos, Y_pos, L, H, Text_size = self.translate_pos(0.3, 0.75, 0.35, 0.1, 0.015)
        self.mainWindowBtn_excelToTable.setFixedSize(L, H)
        self.mainWindowBtn_excelToTable.move(X_pos, Y_pos)
        self.mainWindowBtn_excelToTable.setFont(QFont('Times', Text_size))
        self.mainWindowBtn_excelToTable.setText("EXCEL DATA")
        self.mainWindowBtn_excelToTable.clicked.connect(self.LoadExcel)

    def LoadExcel(self):
        self.w = ExcelToTable()
        self.w.show()


    def EPFEntered(self):
        global EPF
        if self.mainWindowTextBox.text() == "":
            self.mainWindowBtn_createNewEntry.setEnabled(False)
            self.mainWindowBtn_scanALabel.setEnabled(False)
        else:
            self.mainWindowBtn_createNewEntry.setEnabled(True)
            self.mainWindowBtn_scanALabel.setEnabled(True)
            EPF = self.mainWindowTextBox.text()

    def translate_pos(self, ratioX, ratioY, ratioLength, ratioHeight, ratioTextSize):
        requiredX = np.int0(ratioX*self.width)
        requiredY = np.int0(ratioY*self.height)
        requiredLength = np.int0(ratioLength*self.width)
        requiredHeight = np.int0(ratioHeight*self.height)
        requiredTextSize = np.int0(ratioTextSize*self.width)
        return requiredX, requiredY, requiredLength, requiredHeight, requiredTextSize
    def createNewEntry(self):
        print('button to create new entry pressed')
        self.w = CreateANewEntryWindow()
        self.w.show()


    def scanALabel(self):
        print('button to scan a label pressed')
        self.w = ScanLabelBarcode()
        self.w.show()


def window():
    app = QApplication(sys.argv)
    win = MyWindow()
    win.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    load_information()
    primaryFunctions()
    window()