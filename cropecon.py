"""
Need to use xlrd and xlwt to read and write Microsoft Excel Files
And xlutils for some utils


"""

import xlrd, xlwt
import xlutils 
import csv

#opens the workbook
print("Opening Workbook")
wb = xlrd.open_workbook("ag_hr_1998.xls")

#opens the worksheet: can do this by index or name 
print("Opening Worksheets")
icahr = wb.sheet_by_name('ICA HR') # might need to escape space
awhr = wb.sheet_by_name('AW HR')
etawhr = wb.sheet_by_name('ETAW HR')

#data cell's values are accessible by sheet.cell(row,column).value\
#should print out the first year in both (1998)
#print(icahr.cell(1,0).value) 
#print(etawhr.cell(1,0).value) 
#print(icahr.cell(1,5).value)
#for NC in range (5,24):
#	print(icahr.cell(1,NC).value)

ICA_row = [] #make a list for ica table
AW_row = []
ETAW_row = []

for row in range (1,11):
	for col in range (5,25):				#create list for row for ICA
		#print(icahr.cell(row,col).value), #this prevents columns from printing new lines
		#print(" "),
		ICA_row.append(icahr.cell(row,col).value)
	#print("") #this prints a new line to seperate the rows
	for col in range (3,23):
		AW_row.append(awhr.cell(row,col).value)
	for col in range (3,23):
		ETAW_row.append(etawhr.cell(row,col).value)
	print(ICA_row)
	print(AW_row)
	print(ETAW_row)
	ICA_row[:] = []
	AW_row[:] = []
	ETAW_row[:] = []
	print("")
	
#we can now iterate through the list and make each row a list 


#then we make a dot product between the two lists





"""
References: 
http://www.sitepoint.com/using-python-parse-spreadsheet-data/ 
http://stackoverflow.com/questions/4093989/dot-product-in-python

"""