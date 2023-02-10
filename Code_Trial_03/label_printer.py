import win32print
import win32ui
from PIL import Image, ImageWin
import os
import sys
from label import Label

def Print_Labels(Barcode_Info, Mat_Description,twisting,Lot_Number,Machine_Code,BoxQuantity, Pallets, Invoice_Num,Pellet_Num,Customers,Supplier_Name):
    # Label Printer Object
    label_printer = LabelPrinter()
    label1 = Label('Required_Files/information.json')  # Provide the Json file for the label
    # Load values from Json file
    label1.load_variables()
    # Create the barcode
    barcode = label1.create_barcode(Barcode_Info)  # Provide Information to be included in barcode
    # Create the label
    label1.save_label_design(barcode, Mat_Description, twisting, str(Lot_Number), str(Machine_Code), str(BoxQuantity),str(Pallets), Invoice_Num,str(Pellet_Num),Customers,Supplier_Name)  # Provide a Name to save the label

    # Print the label
    label_printer.print_label(label1, 1)
    #delete the temp file
    os.remove('label.png')


class LabelPrinter:

    def __init__(self):
        self.printer_name = win32print.GetDefaultPrinter()

    def print_label(self, label, number_of_prints):

        # Printing scale
        down_scale = round(1 / label.upscale, 1)
        printer_size = [int(label.mainwindow_w * down_scale), int(label.mainwindow_h * down_scale)]

        file_name = "label.png"

        for num in range(number_of_prints):

            hDC = win32ui.CreateDC()
            hDC.CreatePrinterDC(self.printer_name)

            bmp_opened = Image.open(file_name)
            bmp = bmp_opened.resize((printer_size[0], printer_size[1]))

            if bmp.size[0] < bmp.size[1]:
                # bmp = bmp.rotate(90)
                pass
            hDC.StartDoc(file_name)
            hDC.StartPage()

            dib = ImageWin.Dib(bmp)

            dib.draw(hDC.GetHandleOutput(), (0, 0, printer_size[0], printer_size[1]))

            hDC.EndPage()
            hDC.EndDoc()
            hDC.DeleteDC()





