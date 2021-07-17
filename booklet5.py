

from pytesseract import image_to_string
import pytesseract
from PIL import Image ,ImageOps
import numpy as np 
import cv2 
import image
import imutils
from openpyxl import load_workbook
import xlwt
from datetime import datetime
from xlutils.copy import copy # http://pypi.python.org/pypi/xlutils
from xlrd import open_workbook # http://pypi.python.org/pypi/xlrd
import os
from tkinter import Tk     # from tkinter import Tk for Python 3.x
from tkinter.filedialog import askdirectory 



global regno

regno =9


global yi
global yf

global xi
global xf



def excelInsert(j,IE,IEE) :

		global regno


		file_path='black.xls'

		rb = open_workbook("black.xls")
		wb = copy(rb)

		ws = wb.get_sheet(0)
		ws.write(regno, j+1,IE)
		ws.write(regno+1, j+1,IEE)
		wb.save('black.xls')
		# style0 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on',
		#     num_format_str='#,##0.00')
		# style1 = xlwt.easyxf(num_format_str='D-MMM-YY')

		# rb = open_workbook(file_path,formatting_info=True)
		# r_sheet = rb.sheet_by_index(0) # read only copy to introspect the file
		# wb = copy(rb) # a writable copy (I can't read values out of this, only write to it)
		# ws = wb.get_sheet(0) # the sheet to write to within the writable copy

		# wb = xlwt.Workbook()
		# ws = wb.add_sheet('A Test Sheet')

		

		# print('The data '+IE + IEE +"The J"+str(j+2))
		# wb.save('black.xls')

		# wb.save(file_path + '.out' + os.path.splitext(file_path)[-1])

	



def readValue(image,xi,xf,yi,yf):

		image = cv2.imread(image,0)
		image=image[yi:yf, xi:xf]

		# image = imutils.resize(image, width=400)
		blur = cv2.GaussianBlur(image, (7,7), 0)
		thresh = cv2.threshold(blur, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]
		result = 255 - thresh 


		# data = pytesseract.image_to_string(result, lang='eng',config='--psm 6')

		data = pytesseract.image_to_string(image, lang='eng', 
	        config='--psm 10 --oem 3 -c tessedit_char_whitelist=0123456789')

		# if data is None:
		# 	data="null"

		data =data.replace("\u000C",'')
		# data =data.replace(' ','')
		data=data. rstrip("\n")
		data=data. rstrip('     ')
		print(data)

#view & unview the conversion process
		cv2.imshow('thresh', thresh)
		cv2.imshow('result', result)
		cv2.waitKey()

		return data








def readDocument(filename):



		global yi
		global yf

		global xi
		global xf

		yi=650
		yf=809
		xi=1840
		xf=1989

		image =Image.open(filename).convert('RGB')

		box = (1229,159 , 2343,257)
		image_reg = image.crop(box)
		image_reg.save('reg.png')

		Reg_number = pytesseract.image_to_string(image_reg)
		Reg_number=Reg_number.replace("Registration number: ",'')
		Reg_number =Reg_number.replace("\u000C",'')
		Reg_number=Reg_number.rstrip("\n")
		print(Reg_number)

		i=1
		j=1

		file_path='black.xls'

		rb = open_workbook("black.xls")
		wb = copy(rb)

		ws = wb.get_sheet(0)
		ws.write(regno, 0,Reg_number)
		wb.save('black.xls')





		while(i<=6):

			print('Q'+str(i))

			while (j<=3):


				

				if (j==1):
				   QN=readValue(filename,xi,xf,yi,yf)

				elif j==2 :
				   IE=readValue(filename,xi,xf,yi,yf)
				else:
				   IEE=readValue(filename,xi,xf,yi,yf)





				xi=xf+17
				xf=xi+148

				j=j+1

			excelInsert(i,IE,IEE)

			j=1
			i=i+1
			yi=yf+14
			yf=yi+157
			xi=1840
			xf=2000


def openFolder(folderpath):

	global regno

	for  root,dirs,files in os.walk(folderpath): #Lisiting all files
		for entry in files:
			if entry.endswith(('.png', '.jpg', '.jpeg')):
			   #comparing searched file with available files
				# entry='"' +entry+'"'
				#return os.path.join(root,entry)
				a=os.path.join(root,entry) #Getting a full path of the file
				print(a) 
				readDocument(a)

				regno=regno+2




				# os.system(a) #Opening the file
	   




Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
folderpath = askdirectory () # show an "Open" dialog box and return the path to the selected file
print(folderpath)
   
openFolder(folderpath)


# cv2.destroyAllWindows()



