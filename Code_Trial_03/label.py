import json
from PIL import Image, ImageFont, ImageDraw
import os
from barcode import *
from barcode.writer import ImageWriter


class Label:
    def __init__(self, json_file):
        self.json_file = json_file
        self.mainwindow_w, self.mainwindow_h, self.barcode_w, self.barcode_h, self.supplier_x, self.supplier_y, self.mat_des_x, self.mat_des_y, self.mat_code_x, self.mat_code_y, self.lot_num_x, self.lot_num_y, self.num_board_x, self.num_board_y, self.qty_x, self.qty_y, self.invent_num_x, self.invent_num_y, self.fnt_size = 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 12
        self.label_name = "sample_label"

        self.upscale = 8  # set up scale value for creating label

    def load_variables(self):
        info_json_file = open(self.json_file)
        label_info = json.load(info_json_file)

        # Scaling up label image to 10 times
        self.mainwindow_w = int(label_info["mainwindow"]["Total_width_of_the_Label"] * self.upscale)
        self.mainwindow_h = int(label_info["mainwindow"]["Total_height_of_the_label"] * self.upscale)
        self.y_Start = label_info["mainwindow"]["Starting_Y"]

        self.barcode_h = label_info["barcode"]["Total_height_of_the_barcode_area"]
        self.barcode_w = label_info["barcode"]["Total_width_of_the_barcode_area"]

        self.fnt_size_Normal = label_info["FontSelection"]["Font_Size_Normal_Text"]
        self.fnt_size_answer = label_info["FontSelection"]["Font_Size_Ans"]
        self.fnt_size_first = label_info["FontSelection"]["First_two_fonts"]

        self.gap_x = label_info["text_partition"]["X_position_of_gap"]
        self.gap_y = label_info["text_partition"]["Y_position_of_gap"]

        self.mat_description_x = label_info["text_partition"]["X_position_of_mat_description"]
        self.mat_description_y = label_info["text_partition"]["Y_position_of_mat_description"]

    def create_barcode(self, barcode_info):
        # Name for the barcode
        barcode_name = barcode_info
        # pass the information to be include in the barcode with the ImageWriter() as the writer
        # barcode = EAN13(barcode_info, writer=ImageWriter())
        image_writer = ImageWriter()
        image_writer.font_path = 'DejaVuSansMono.ttf'
        barcode = Code128(barcode_info, writer=image_writer)
        # Save Barcode
        PATH = os.getcwd()
        name = PATH+"\\" + barcode_info
        barcode.save(name)
        return barcode_name

    def save_label_design(self,barcode_name, Mat_Description,twisting,Lot_Number,Macine_Code,BoxQuantity, Pallets, Invoice_Num,Pellet_Num,Customers,Supplier_Name):
        # Create a empty white layout for the label
        layout = Image.new('RGB', (self.mainwindow_w, self.mainwindow_h), (255, 255, 255))

        # Add text info to the label

        title_text_1 = "Material Description:"
        title_text_2 = "Twist (S/Z):"
        title_text_3 = "Lot Number:"
        title_text_4 = "Material Code:"
        title_text_5 = "Box Quantity:"
        title_text_6 = "Kg in Pellet:"
        title_text_7 = "Invoice:"
        title_text_8 = "Pellet Number:"
        title_text_9 = "Customer:"
        title_text_10 = "Supplier:"

        image_editable = ImageDraw.Draw(layout)

        gap = self.gap_y * self.mainwindow_h


        texts_font = ImageFont.truetype('arial', int(self.fnt_size_Normal * self.mainwindow_h))
        ans_font = ImageFont.truetype('arial', int(self.fnt_size_answer * self.mainwindow_h))
        w, h = image_editable.textsize(title_text_1, font=texts_font)
        x_in_frist_column =self.gap_x+h/2

        texts_font_first = ImageFont.truetype('arial', int(self.fnt_size_first * self.mainwindow_h))

        ### Numbers font size
        ans_font_Numbers = ImageFont.truetype('arial', int(self.fnt_size_answer * self.mainwindow_h*1.2))
        ### Material description font size
        texts_font_first_MD = texts_font_first
        if len(Mat_Description) > 30:
            texts_font_first_MD = ImageFont.truetype('arial', int(self.fnt_size_first * self.mainwindow_h *0.3))
        elif len(Mat_Description) >= 20 and len(Mat_Description) < 30:
            texts_font_first_MD = ImageFont.truetype('arial', int(self.fnt_size_first * self.mainwindow_h * 0.5))
        elif len(Mat_Description) >= 11 and len(Mat_Description) < 20:
            texts_font_first_MD = ImageFont.truetype('arial', int(self.fnt_size_first * self.mainwindow_h * 0.7))
        ### Lot number font size
        ans_font_LN = ans_font
        if len(Lot_Number) > 30:
            ans_font_LN = ImageFont.truetype('arial', int(self.fnt_size_answer * self.mainwindow_h *0.3))
        elif len(Lot_Number) >= 20 and len(Lot_Number) < 30:
            ans_font_LN = ImageFont.truetype('arial', int(self.fnt_size_answer * self.mainwindow_h * 0.5))
        elif len(Lot_Number) >= 11 and len(Lot_Number) < 20:
            ans_font_LN = ImageFont.truetype('arial', int(self.fnt_size_answer * self.mainwindow_h * 0.7))
        ### Material code font size
        ans_font_MC = ans_font
        if len(Macine_Code) > 30:
            ans_font_MC = ImageFont.truetype('arial', int(self.fnt_size_answer * self.mainwindow_h * 0.3))
        elif len(Macine_Code) >= 20 and len(Macine_Code) < 30:
            ans_font_MC = ImageFont.truetype('arial', int(self.fnt_size_answer * self.mainwindow_h * 0.5))
        elif len(Macine_Code) >= 11 and len(Macine_Code) < 20:
            ans_font_MC = ImageFont.truetype('arial', int(self.fnt_size_answer * self.mainwindow_h * 0.7))

        ### Invoice font size
        ans_font_In = ans_font
        if len(Invoice_Num) > 30:
            ans_font_In = ImageFont.truetype('arial', int(self.fnt_size_answer * self.mainwindow_h * 0.3))
        elif len(Invoice_Num) >= 20 and len(Invoice_Num) < 30:
            ans_font_In = ImageFont.truetype('arial', int(self.fnt_size_answer * self.mainwindow_h * 0.5))
        elif len(Invoice_Num) >= 11 and len(Invoice_Num) < 20:
            ans_font_In = ImageFont.truetype('arial', int(self.fnt_size_answer * self.mainwindow_h * 0.7))

        ### Customer font size
        ans_font_Cus = ans_font
        if len(Customers) > 30:
            ans_font_MC = ImageFont.truetype('arial', int(self.fnt_size_answer * self.mainwindow_h * 0.3))
        elif len(Customers) >= 20 and len(Customers) < 30:
            ans_font_MC = ImageFont.truetype('arial', int(self.fnt_size_answer * self.mainwindow_h * 0.5))
        elif len(Customers) >= 11 and len(Customers) < 20:
            ans_font_MC = ImageFont.truetype('arial', int(self.fnt_size_answer * self.mainwindow_h * 0.7))
        ### Supplier font size
        ans_font_Sup = ans_font
        if len(Supplier_Name) > 30:
            ans_font_Sup = ImageFont.truetype('arial', int(self.fnt_size_answer * self.mainwindow_h * 0.3))
        elif len(Supplier_Name) >= 20 and len(Supplier_Name) < 30:
            ans_font_Sup = ImageFont.truetype('arial', int(self.fnt_size_answer * self.mainwindow_h * 0.5))
        elif len(Supplier_Name) >= 11 and len(Supplier_Name) < 20:
            ans_font_Sup = ImageFont.truetype('arial', int(self.fnt_size_answer * self.mainwindow_h * 0.7))



        image_editable.text((x_in_frist_column, gap), title_text_1, (0, 0, 0), font=texts_font)

        image_editable.text((x_in_frist_column, 3.5*gap ), Mat_Description, (65, 65, 65),font=texts_font_first_MD, stroke_width=2)

        image_editable.text((self.mainwindow_w / 2+ 5.5*gap , gap), title_text_2, (0, 0, 0), font=texts_font, stroke_width=1)

        image_editable.text((self.mainwindow_w / 2 + 10 *gap, 3.5 * gap), twisting,(65, 65, 65),font=texts_font_first, stroke_width=2)

        #---------------------------------------------------------------------------------------------------------------------

        H_Previous_matdes = gap + 4 * h

        image_editable.text((x_in_frist_column, H_Previous_matdes + gap), title_text_3, (0, 0, 0), font=texts_font)

        image_editable.text((x_in_frist_column, 3*gap +H_Previous_matdes + gap), Lot_Number, (65, 65, 65), font=ans_font_LN, stroke_width=2)

        image_editable.text((self.mainwindow_w / 2 + gap / 2, H_Previous_matdes + gap), title_text_4, (0, 0, 0), font=texts_font)

        image_editable.text((self.mainwindow_w / 2 + gap/2, 3*gap+H_Previous_matdes + gap), Macine_Code, (65, 65, 65), font=ans_font_MC,stroke_width=2)

        # ---------------------------------------------------------------------------------------------------------------------

        H_Previous_lot = 2*(gap + 4 * h)

        image_editable.text((x_in_frist_column,H_Previous_lot+gap), title_text_5, (0, 0, 0), font=texts_font)

        image_editable.text((x_in_frist_column, 3 * gap + H_Previous_lot + gap), BoxQuantity, (65, 65, 65), font=ans_font_Numbers, stroke_width=2)

        image_editable.text((self.mainwindow_w / 2 + gap / 2, H_Previous_lot + gap), title_text_6, (0, 0, 0), font=texts_font,stroke_width=1)

        image_editable.text((self.mainwindow_w / 2 + gap / 2, 3 * gap + H_Previous_lot + gap), Pallets,(65, 65, 65), font=ans_font_Numbers, stroke_width=2)

        # ---------------------------------------------------------------------------------------------------------------------

        H_Previous_BoxQ = 3 * (gap + 4 * h)

        image_editable.text((x_in_frist_column, H_Previous_BoxQ + gap), title_text_7, (0, 0, 0), font=texts_font)

        image_editable.text((x_in_frist_column, 3 * gap + H_Previous_BoxQ + gap), Invoice_Num, (65, 65, 65), font=ans_font_In, stroke_width=2)

        image_editable.text((self.mainwindow_w / 2 + gap / 2, H_Previous_BoxQ + gap), title_text_8, (0, 0, 0), font=texts_font, stroke_width=1)

        image_editable.text((self.mainwindow_w / 2 + gap / 2, 3 * gap + H_Previous_BoxQ + gap), Pellet_Num, (65, 65, 65), font=ans_font_Numbers, stroke_width=2)

        # ---------------------------------------------------------------------------------------------------------------------

        H_Previous_Invoice = 4 * (gap + 4 * h)

        image_editable.text((x_in_frist_column, H_Previous_Invoice + gap), title_text_9, (0, 0, 0), font=texts_font)

        image_editable.text((x_in_frist_column, 3 * gap + H_Previous_Invoice + gap), Customers, (65, 65, 65), font=ans_font_Cus, stroke_width=2)

        image_editable.text((self.mainwindow_w / 2 + gap / 2, H_Previous_Invoice + gap), title_text_10, (0, 0, 0), font=texts_font, stroke_width=1)

        image_editable.text((self.mainwindow_w / 2 + gap / 2, 3 * gap + H_Previous_Invoice + gap), Supplier_Name, (65, 65, 65), font=ans_font_Sup, stroke_width=2)

        H_Previous_Supplier = 5 * (gap + 4 * h)

        # ---------------------------------------------------------------------------------------------------------------------


        image_editable.line([(0, gap+4*h), (self.mainwindow_w, gap+4*h)], fill="black", width=int(0.001428571429 * self.mainwindow_h))
        image_editable.line([(0, (gap + 4 * h)*2), (self.mainwindow_w, (gap + 4 * h)*2)], fill="black",width=int(0.001428571429 * self.mainwindow_h))
        image_editable.line([(0, (gap + 4 * h) * 3), (self.mainwindow_w, (gap + 4 * h) * 3)], fill="black", width=int(0.001428571429 * self.mainwindow_h))
        image_editable.line([(0, (gap + 4 * h) * 4), (self.mainwindow_w, (gap + 4 * h) * 4)], fill="black", width=int(0.001428571429 * self.mainwindow_h))
        image_editable.line([(0, (gap + 4 * h) * 5), (self.mainwindow_w, (gap + 4 * h) * 5)], fill="black",width=int(0.001428571429 * self.mainwindow_h))

        image_editable.line([(self.mainwindow_w / 2, gap+4*h), (self.mainwindow_w/2, (gap + 4 * h) * 5)], fill="black", width=int(0.001428571429 * self.mainwindow_h))
        image_editable.line([(self.mainwindow_w / 2+ 5*gap,0), (self.mainwindow_w / 2+ 5*gap, gap + 4 * h)], fill="black", width=int(0.001428571429 * self.mainwindow_h))

        #----------------------------------------------------------------------------------------------

        line_y_pos = ((gap + 4 * h) * 5)
        # print("last line", line_y_pos)

        # Open created barcode
        barcode = Image.open(barcode_name + ".png")

        actual_barcode_w, actual_barcode_h = barcode.size

        # print("actual_barcode_size", actual_barcode_w, actual_barcode_h)
        # print("possible_space", self.mainwindow_w, self.mainwindow_h - line_y_pos)

        for scale in range(8, 1, -1):
            scale = 0.1 * scale
            # new_barcode_w, new_barcode_h = scale * self.mainwindow_w, scale * (self.mainwindow_h - line_y_pos) #same scale for width & height
            new_barcode_w, new_barcode_h = scale * self.mainwindow_w, (
                    scale * self.mainwindow_w * actual_barcode_h) / actual_barcode_w
            # print("new_barcode_size", scale, new_barcode_w, new_barcode_h)
            if new_barcode_w < self.mainwindow_w and new_barcode_h < (self.mainwindow_h - line_y_pos):
                barcode_pos_x = int((self.mainwindow_w - new_barcode_w) / 2)
                barcode_pos_y = int((self.mainwindow_h + line_y_pos - new_barcode_h) / 2)
                resized_barcode = barcode.resize((int(new_barcode_w), int(new_barcode_h)))
                break

        # Place the barcode on the layout
        layout.paste(resized_barcode, (barcode_pos_x, barcode_pos_y))

        #layout.show()  #mark this #

        layout.save("label.png")
        nameToDelete = str(barcode_name) + ".png"   #unmark this #
        os.remove(nameToDelete)                         #unmark this #

