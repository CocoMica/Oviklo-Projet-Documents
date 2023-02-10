import os
directory = os.fsencode("Backup_Files")

def readFileNames():
    for file in os.listdir(directory):
        filename = os.fsdecode(file)
        print(filename)



def readFiles():
    with open('readme.txt', 'w') as fileWrite:
        Headings = 'Date Stamp\tTime Stamp\tType Of Record\tData Entered By\tOverall Barcode\tPO Number\tExcel 01 Row\tPellet Number\tLabel issue number\tSupplier\tCustomer\tMaterial Description\tMaterial Code\tTwist\tLot Number\tInvoice Number\tTotal Initial Box Qty\tBox Qty In Pellet Before Scanning\tBox Quantity Removed\tNew Box Qty In Pellet\tTotal Initial Weight\tWeight In Pellet Before Scanning\tWeight Removed\tNew Weight In Pellet\n'
        fileWrite.write(Headings)

        for file in os.listdir(directory):
            filename = os.fsdecode(file)
            print("____________________________NEW FILE________________________________")
            print(filename)
            fname = "Backup_Files/"+filename
            with open(fname) as f:
                for line in f:
                    nanFound = line.find('nan')
                    if nanFound ==-1:
                        print(line.strip())
                        fileWrite.write(line.strip())
                        fileWrite.write('\n')



# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    readFiles()
    #readFileNames()



