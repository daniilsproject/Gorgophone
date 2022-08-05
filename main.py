import os
from win32com import client
from random import choice
import string
from datetime import date
from PyPDF2 import PdfMerger
import config
import delete


today = date.today()
# Month abbreviation, day and year	
d3 = today.strftime("%d%m%y")
print("Today is ", d3)


#get paths
path_to_inputs = config.path_to_inputs
path_to_save = config.path_to_save


#metod to generte a random name
def id_generator(size=8, chars=string.ascii_uppercase + string.digits):
    return ''.join(choice(chars) for _ in range(size))


#set up application
app = client.DispatchEx("Excel.Application")
app.Interactive = False
app.Visible = False
pdf_files = []


for filename in os.listdir(path = path_to_inputs):

    #excluding temporary file
    if ((filename[0] != '~') and (filename[1] != '$')):
        
        print ("Open file: "+filename)
        
        #give your file name with valid path
        input_file = path_to_inputs + '/' +filename
        
        #open file
        Workbook = app.Workbooks.Open(input_file)

        #get & save sheets names
        for sheet in Workbook.Sheets:
            pages_name=path_to_save + "/" +id_generator()+'.pdf'
            print(" Page: "+ sheet.Name +" saved: "+pages_name)
            pdf_files.append(pages_name)
            #output_file = r'F:\_work\LX\pdfConverter\result\'+pages_name
            try:
                Workbook.Worksheets(sheet.name).ExportAsFixedFormat(0, pages_name)
                print(" Done")
            except Exception as e:
                print("Failed to convert in PDF format.Please confirm environment meets all the requirements  and try again")
                print(str(e))

        Workbook.Close()




print()
print("Starting mergin...")
merger = PdfMerger()


for pdf in pdf_files:
    merger.append(pdf)


merger.write(path_to_save + "/" +"__"+d3+'_merged'+'.pdf')
print(" merged file path:"+path_to_save + "/" +d3+'_merged'+'.pdf')
merger.close()

#deleting
print("Do you want to delete all one-page files?")
print("Type \'Y\' or \'y\' if you want to or any key to not to")
delete_request = input()

if ((delete_request == "Y") or (delete_request == "y")):
    delete.remove_one_page_files(path_to_save)


app.Exit()
