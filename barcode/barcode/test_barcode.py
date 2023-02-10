import json
import win32print
from win32printing import Printer
import win32ui
from PIL import ImageWin
from PIL import Image, ImageFont, ImageDraw

from barcode import EAN13, ISBN10, Code39, PZN, Code128

from barcode.writer import ImageWriter

Info_JSON_Path = 'test1.json'

printer_name = win32print.GetDefaultPrinter()


class BarcodePrinter:

    def Load_Variables(self):

        global mainwindow_w, mainwindow_h, barcode_w, barcode_h, fnt_size, supplier_x, supplier_y, mat_des_x, mat_des_y, mat_code_x, mat_code_y, lot_num_x, lot_num_y, num_board_x, num_board_y, qty_x, qty_y, invent_num_x, invent_num_y

        Info_Json_Loc = open(Info_JSON_Path)
        Info_Total = json.load(Info_Json_Loc)
        mainwindow_w = Info_Total["mainwindow"]["Total_width_of_the_Label"]
        mainwindow_h = Info_Total["mainwindow"]["Total_height_of_the_label"]

        barcode_w = Info_Total["barcode"]["Total_height_of_the_barcode_area"]
        barcode_h = Info_Total["barcode"]["Total_width_of_the_barcode_area"]

        fnt_size = Info_Total["text_partition"]["Font_Size"]
        supplier_x = Info_Total["text_partition"]["X_position_of_supplier"]
        supplier_y = Info_Total["text_partition"]["Y_position_of_supplier"]
        mat_des_x = Info_Total["text_partition"]["X_position_of_material_description"]
        mat_des_y = Info_Total["text_partition"]["Y_position_of_material_description"]
        mat_code_x = Info_Total["text_partition"]["X_position_of_material_code"]
        mat_code_y = Info_Total["text_partition"]["Y_position_of_material_code"]
        lot_num_x = Info_Total["text_partition"]["X_position_of_lot_number"]
        lot_num_y = Info_Total["text_partition"]["Y_position_of_lot_number"]
        num_board_x = Info_Total["text_partition"]["X_position_of_number_of_boards"]
        num_board_y = Info_Total["text_partition"]["Y_position_of_number_of_boards"]
        qty_x = Info_Total["text_partition"]["X_position_of_qty"]
        qty_y = Info_Total["text_partition"]["Y_position_of_qty"]
        invent_num_x = Info_Total["text_partition"]["X_position_of_inventory_number"]
        invent_num_y = Info_Total["text_partition"]["Y_position_of_inventory_number"]

        print("Assignment Complete")

    def print_barcode(self, label_name, number_of_prints):

        printer_size = [mainwindow_w, mainwindow_h]

        file_name = label_name + ".png"

        for num in range(number_of_prints):

            hDC = win32ui.CreateDC()
            hDC.CreatePrinterDC(printer_name)

            #printer_size = [hDC.GetDeviceCaps(PHYSICALWIDTH), hDC.GetDeviceCaps(PHYSICALHEIGHT)]


            bmp = Image.open(file_name)



            if bmp.size[0] < bmp.size[1]:
                # bmp = bmp.rotate(90)
                pass

            hDC.StartDoc(file_name)
            hDC.StartPage()

            dib = ImageWin.Dib(bmp)

            #dib.draw(hDC.GetHandleOutput(), (0, 0, int(0.4 * printer_size[0]), int(0.1 * printer_size[1])))
            dib.draw(hDC.GetHandleOutput(), (0, 0, int(printer_size[0]), int(printer_size[1])))

            hDC.EndPage()
            hDC.EndDoc()
            hDC.DeleteDC()

            print("Complete " + str(num))

    def create_barcode(self, barcode_info):

        # pass the information to be include in the barcode with the ImageWriter() as the writer
        # my_code = EAN13(number, writer=ImageWriter())
        barcode = Code128(barcode_info, writer=ImageWriter())

        # Save Barcode
        barcode.save(barcode_info)

    def saveLabelDesign(self, barcode_name, label_name):

        # Create a empty white layout for the label
        layout = Image.new('RGB', (mainwindow_w, mainwindow_h), (255, 255, 255))

        # Open created barcode
        barcode = Image.open(barcode_name + ".png")
        # Place the barcode on the layout
        layout.paste(barcode, (barcode_w, barcode_h))

        # Add text info to the label
        title_font = ImageFont.truetype('arial', 40)
        texts_font = ImageFont.truetype('arial', fnt_size)

        title_text = "Description"

        title_text_1 = "Supplier "
        title_text_2 = "Material Description "
        title_text_3 = "Material Code "
        title_text_4 = "Lot Number "
        title_text_5 = "Number of Blocks "
        title_text_6 = "Qty (kg) "
        title_text_7 = "Invoice Number "

        image_editable = ImageDraw.Draw(layout)

        image_editable.text((150, 0), title_text, (0, 0, 0), font=title_font)

        image_editable.text((supplier_x, supplier_y), title_text_1, (0, 0, 0), font=texts_font)
        image_editable.text((mat_des_x, mat_des_y), title_text_2, (0, 0, 0), font=texts_font)
        image_editable.text((mat_code_x, mat_code_y), title_text_3, (0, 0, 0), font=texts_font)
        image_editable.text((lot_num_x, lot_num_y), title_text_4, (0, 0, 0), font=texts_font)
        image_editable.text((num_board_x, num_board_y), title_text_5, (0, 0, 0), font=texts_font)
        image_editable.text((qty_x, qty_y), title_text_6, (0, 0, 0), font=texts_font)
        image_editable.text((invent_num_x, invent_num_y), title_text_7, (0, 0, 0), font=texts_font)

        layout.save(label_name + ".png")


# Barcode creator object
barcode_printer_obj = BarcodePrinter()
# Load values from Json file
barcode_printer_obj.Load_Variables()
# Create the barcode : data to be in barcode
barcode_printer_obj.create_barcode("5000005852")
# Create the label : (barcode name, label name to be)
barcode_printer_obj.saveLabelDesign("5000005852", "5000005852_label")
# Print the label : label_name
barcode_printer_obj.print_barcode("5000005852_label", 1)
