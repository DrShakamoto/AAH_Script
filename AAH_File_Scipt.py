import openpyxl
import datetime
import os
from win32com.client import Dispatch
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import zipfile

filename = "STKPREAN.FLT"

def file_to_list(filename):#function to convert to .zip and put each line from the file into a list

	if os.path.isfile('PREANZIP.txt'):
		os.rename('PREANZIP.txt', 'PREAN.zip')#rename the txt file to the zip
		with zipfile.ZipFile("PREAN.zip","r") as zip_ref:#unzip the file
			zip_ref.extractall()
		os.remove('PREAN.zip')#remove the zip file

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
	
	workbook.save(newFilename)#saves the workbook under a new name

def run_macro(filename):#run the excel macro
	xlApp = Dispatch("Excel.Application")
	cwd = os.getcwd()
	xlApp.Workbooks.Open(cwd + '\\' + filename)
	xlApp.Visible=0
	xlApp.Run("AAHFileMacro")
	xlApp.Workbooks(1).Close(SaveChanges=1)
	xlApp.Application.Quit()

def send_email(filename):
	sendEmail = input("Would you like to send the spreadsheet to hwright@weldricks.co.uk? 'y' or 'n': ")#check if the user would like to send the email
	
	if sendEmail == 'y' or sendEmail == 'Y':

		s = smtplib.SMTP(host='mail.weldricks.co.uk', port=25)#set the mail server and port
		s.starttls()#start the SMTP session

		msg = MIMEMultipart() #create a message

		#set the parameters of the message:
		msg['From'] = ""
		msg['To'] = ""
		msg['Cc'] = ""
		msg['Bcc'] = ""
		msg['Subject'] = "AAH Product File"

		body = "Hi Harry,\n\nPlease find the weekly AAH product spreadsheet attached.\n\nThis message was automatically generated."
		msg.attach(MIMEText(body, 'plain'))

		attachment = open(filename, 'rb')

		part = MIMEBase('application', 'octet-stream')
		part.set_payload((attachment).read())
		encoders.encode_base64(part)
		part.add_header('Content-Disposition', "attachment; filename= %s" %filename)

		msg.attach(part)

		s.send_message(msg)#send the message
		del msg#delete local copy of message
		s.quit()#close the SMTP session


	else:
		input("The spreadsheet has been generated but not sent. Press any key to exit.")

current_date=datetime.datetime.today().strftime('%Y%d%m')#get current date for file name
newFilename = current_date + '.xlsm'
newFilename = newFilename.replace('-', '')#create the new filename in format yyyyddmm

columns_into_excel(line_to_columns(file_to_list(filename)))
run_macro(newFilename)
send_email(newFilename)
