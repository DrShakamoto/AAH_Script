import openpyxl
import datetime
import os
from win32com.client import Dispatch

filename = "STKPREAN.FLT"

def file_to_list(filename):#function to place each line from the file into a list
	file = open(filename,"r") #open the file as read only

	line_list = []#creating a blank list to store each line in 

	for line in file:
		line_list.append(line)#adds each line into the list

	file.close()#close the file to free up resources

	return line_list#function returns the list of lines

def line_to_columns(line_list):#function to seperate each line into the correctly sized columns
	for i in range(0, len(line_list)):
		string_list = list(line_list[i])#creates a list of every character in the line

		"""
		add commas at the correct places in the string to use as delimiters:

		"""
		string_list.insert(1, ',')
		string_list.insert(42, ',')
		string_list.insert(51, ',')
		string_list.insert(60, ',')
		string_list.insert(69, ',')
		string_list.insert(79, ',')
		string_list.insert(89, ',')
		string_list.insert(94, ',')
		"""
		remove the trailing \n from the string:

		"""
		string_list.pop()
		string_list.pop()

		
		string = ''.join(string_list)#joins the list back into a string including added commas
		column_list = string.split(",")#creates a list containing each column
		line_list[i] = column_list#replaces the original string stored in the line_list with a second list containing each column to create a 2d list
	
	return line_list

def columns_into_excel(line_list):#function to put the data into excel
	workbook = openpyxl.load_workbook(filename='Macro_Template.xlsm', read_only=False, keep_vba=True)#open macro template
	sheet = workbook.active#set it to use the only sheet in the workbook

	Columns = ['A','B','C','D','E','F','G','H','I']

	for i in range(0, len(Columns)):#loops once for every column
		column = Columns[i]#gets the colum letter from the columns list
		for z in range(0, len(line_list)):#loops once for every row
			row = str(z+1)#gets the row number
			cell = column + row#stores the cell value as a string

			try:#if the value is a number, convert it to an integer. If not, continue with it as a string. This prevents the "number stored as text" message in excel
				line_list[z][i] = int(line_list[z][i])
			except:
				pass
			
			sheet[cell] = line_list[z][i]#sets the value of each cell

	#current_date=datetime.datetime.today().strftime('%Y%d%m')#get current date for file name
	#newFilename = current_date + '.xlsm'
	#newFilename = newFilename.replace('-', '')#create the new filename which is the date in format yyyyddmm
	
	workbook.save(newFilename)#saves the workbook under a new name

def run_macro(filename):
	xlApp = Dispatch("Excel.Application")
	cwd = os.getcwd()
	xlApp.Workbooks.Open(cwd + '\\' + filename)
	xlApp.Visible=1
	xlApp.Run("AAHFileMacro")

current_date=datetime.datetime.today().strftime('%Y%d%m')#get current date for file name
newFilename = current_date + '.xlsm'
newFilename = newFilename.replace('-', '')#create the new filename which is the date in format yyyyddmm

columns_into_excel(line_to_columns(file_to_list(filename)))
run_macro(newFilename)