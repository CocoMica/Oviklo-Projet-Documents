from datetime import datetime
global Backup_File_Name


def Generate_Backup_Data(path):
    global Backup_File_Name
    U_id = str(datetime.today().year)+"-"+str(datetime.today().month)+"-"+str(datetime.today().day) + \
        "-"+str(datetime.today().hour)+"-" + \
        str(datetime.today().minute)+"-"+str(datetime.today().second)
    Backup_File_Name = path+U_id+".txt"


def Write_To_Backup_text_Doc(TypeOfRecord, DataEnteredBy,OverallBarcode, PONumber, Excel01Row, PelletNumber, LabelIssueNumber, Supplier, Customer,  MaterialDescription,MaterialCode, LotNumber, InvoiceNumber, TotalInitialBoxQty,BoxQtyInPelletBeforeScanning, BoxQuantityRemoved, NewBoxQtyInPellet , WeightInPelletBeforeScanning, WeightRemoved, NewWeightInPellet, TotalInitialWeight, Twist):
    global Backup_File_Name
    global Backup_File
    try:
        Backup_File = open(Backup_File_Name, "a+")
        Date = str(datetime.today().year)+"-" + \
            str(datetime.today().month)+"-"+str(datetime.today().day)
        Time = str(datetime.today().hour)+"-" + \
            str(datetime.today().minute)+"-"+str(datetime.today().second)
        Backup_File.write(Date+"\t"+Time+"\t"+str(TypeOfRecord)+"\t"+str(DataEnteredBy)+"\t"+str(OverallBarcode) +"\t"+str(PONumber)+"\t"+str(Excel01Row)+"\t"+str(PelletNumber)+"\t"+str(LabelIssueNumber)+"\t"+str(Supplier)+"\t"+str(Customer)+"\t"+str(MaterialDescription)+"\t"+str(MaterialCode)+"\t"+str(Twist)+"\t"+str(LotNumber)+"\t"+str(InvoiceNumber)+"\t"+str(TotalInitialBoxQty)+"\t"+str(BoxQtyInPelletBeforeScanning)+"\t"+str(BoxQuantityRemoved)+"\t"+str(NewBoxQtyInPellet)+"\t"+str(TotalInitialWeight)+"\t"+str(WeightInPelletBeforeScanning)+"\t"+str(WeightRemoved)+"\t"+str(NewWeightInPellet)+"\n")
        Backup_File.close()
        return True
    except Exception as E:
        print("function Write_To_Backup_text_Doc: Failed to write information to backup file.", E)
        return False
