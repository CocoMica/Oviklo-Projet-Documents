import pandas as pd

def Open_Excel_All_Records(path):
    try:
        df = pd.read_excel(path, sheet_name='Updated')
        return True, df
    except:
        return False, 'Could not connect to the excel. Try closing it.'


def Open_Excel_01_In_UI(path, Reference):
    try:
        Reference = int(Reference)
    except:
        pass
    try:
        df = pd.read_excel(path, sheet_name='Updated')
        df_res = df[df.isin([Reference]).any(1)]
        if len(df_res.index) > 0:
            return True, df_res
        else:
            return False, 'No match for the query'
    except:
        return False, 'Could not connect to the excel. Try closing it.'



def Open_Excel_02(path, Reference):
    try:
        df = pd.read_excel(path, sheet_name='Sheet')
        df_res = df[df.isin([Reference]).any(1)]
        try:
            dateStamp = df_res.iloc[0].at['Date Stamp']
            timeStamp= df_res.iloc[0].at['Time Stamp']
            typeOfRecord= df_res.iloc[0].at['Type Of Record']
            dataEnteredBy= df_res.iloc[0].at['Data Entered By']
            overallBarcode= df_res.iloc[0].at['Overall Barcode']
            PONumber= df_res.iloc[0].at['PO Number']
            Excel01Row= df_res.iloc[0].at['Excel 01 Row']
            pelletNumber= df_res.iloc[0].at['Pellet Number']
            labelIssueNumber= df_res.iloc[0].at['Label issue number']
            supplier = df_res.iloc[0].at['Supplier']
            customer = df_res.iloc[0].at['Customer']
            materialDescription= df_res.iloc[0].at['Material Description']
            materialCode = df_res.iloc[0].at['Material Code']
            twist = df_res.iloc[0].at['Twist']
            lotNumber  = df_res.iloc[0].at['Lot Number']
            invoiceNumber = df_res.iloc[0].at['Invoice Number']
            totalInitialBoxQty = df_res.iloc[0].at['Total Initial Box Qty']
            boxQtyInPelletBeforeScanning = df_res.iloc[0].at['Box Qty In Pellet Before Scanning']
            boxQtyRemoved = df_res.iloc[0].at['Box Quantity Removed']
            newBoxQtyInPellet  = df_res.iloc[0].at['New Box Qty In Pellet']
            totalInitialWeight = df_res.iloc[0].at['Total Initial Weight']
            weightInPelletBeforeScanning = df_res.iloc[0].at['Weight In Pellet Before Scanning']
            weightRemoved = df_res.iloc[0].at['Weight Removed']
            newWeightInPellet = df_res.iloc[0].at['New Weight In Pellet']
            Excel_Row = df_res.index[0]

            response = [dateStamp,timeStamp,typeOfRecord,dataEnteredBy,overallBarcode,PONumber,Excel01Row,pelletNumber,labelIssueNumber,supplier,materialDescription,materialCode,lotNumber,invoiceNumber,totalInitialBoxQty,boxQtyInPelletBeforeScanning, boxQtyRemoved, newBoxQtyInPellet, customer, twist,totalInitialWeight, weightInPelletBeforeScanning, weightRemoved, newWeightInPellet]
            return [response, True, Excel_Row]
        except:
            return [None, False, "No matching reference found for the scanned barcode"]
    except:
        return [None, False, "Please close Excel 02"]

def Check_Excel_02_Record_Duplication(path, Reference):
    #print(path)
    try:
        df = pd.read_excel(path, sheet_name='Sheet')
        df_res = df[df.isin([Reference]).any(1)]
        if len(df_res) == 0:
            return [False,None]
        else:
            return [True, "A label is already created for the following entry."]
    except:
        return [True, "Please close the Excel_02 document."]
