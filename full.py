import openpyxl
import csv
import os
from os import path

data1 = "YOUTUBE_VTUNE"
data2 = "PDF_VTUNE"
def coltocsv(cln):
#This is for Clean Data
	print cln
	allfiles= os.listdir("C:\Users\\\Desktop\DUMMY_DATA2\DUMMY_DATA\\" + cln + "\\")#fetching all the files from directory
	print "Converting Following files"
	print allfiles
	for rownumber in range(3,22): #reading the cells in csv 
		for i in allfiles: #picking files from dir
			
			z = "C:\Users\\\Desktop\DUMMY_DATA2\DUMMY_DATA\\" + cln + "\\" + i
			wb = openpyxl.load_workbook(z)
			sheet = wb.worksheets[0]
			# # #fecthing A B C cells
			x1 = 'A' + str(rownumber)
			y1 = 'B' + str(rownumber)
			z1 = 'C' + str(rownumber)

			# #fetching cell values
			a = sheet[x1].value
			b = sheet[y1].value
			c = sheet[z1].value


			#Writing onto CSV
			dirname = cln +"_column_to_CSV\\" + str(a) + ".csv"
			csvfilename ="C:\Users\\\Desktop\DUMMY_DATA2\DUMMY_DATA\\" + dirname
			file_path = path.relpath(csvfilename)
			with open(file_path, "ab") as csvfile:
			 	writer = csv.writer(csvfile,  quoting=csv.QUOTE_ALL)
			  	writer.writerow([a, b, c, cln, i])

def normalizecolumns(colfolder):
	allfiles= os.listdir("C:\Users\\\Desktop\DUMMY_DATA2\DUMMY_DATA\\"+ colfolder +"_column_to_CSV\\")
	# print allfiles
	for name in allfiles:
		# print name
		filepath = "C:\Users\\\Desktop\DUMMY_DATA2\DUMMY_DATA\\"+ colfolder +"_column_to_CSV\\" + name
		# print filepath

		with open(filepath, 'r') as f:
		 	reader = csv.reader(f)
		 	# print reader
			wholesheetlist = list(reader)
			# print wholesheetlist #whole sheet(single column) in list

		hardwareeventcount_raw = []
		hardwareeventsamplecount_raw = []
		hardwareeventtype = []
		datatype = []
		filename = []

		for i in range (0, len(wholesheetlist)):

			hardwareeventtype.append(wholesheetlist[i][0])
		 	hardwareeventcount = wholesheetlist[i][1] #fetching Hardware Event count
		 	hardwareeventsamplecount = wholesheetlist[i][2]
		 	datatype.append(wholesheetlist[i][3])
		 	filename.append(wholesheetlist[i][4])


		 	hardwareeventcount_raw.append(float(hardwareeventcount))
		 	hardwareeventsamplecount_raw.append(float(hardwareeventsamplecount))

		if ( sum(hardwareeventcount_raw) > 1 ):
			nor_hardwareeventcount = [i/sum(hardwareeventcount_raw) for i in hardwareeventcount_raw]


		
		if ( sum(hardwareeventsamplecount_raw) > 1 ):
			nor_hardwareeventsamplecount = [i/sum(hardwareeventsamplecount_raw) for i in hardwareeventsamplecount_raw]
		
		normalized_directory = "C:\Users\\\Desktop\DUMMY_DATA2\DUMMY_DATA\\"+ colfolder +"_to_CSV_Normalized\\" + name
		

		# print len(hardwareeventtype)
		# print len(nor_hardwareeventcount)
		# print len(nor_hardwareeventsamplecount)
		# print len(datatype)
		# print len(filename)


		with open(normalized_directory, "wb") as csvfile:
		 	writer = csv.writer(csvfile,  quoting=csv.QUOTE_ALL)
		 	for t in xrange(0, len(nor_hardwareeventcount)):
		 		writer.writerow([hardwareeventtype[t], nor_hardwareeventcount[t], nor_hardwareeventsamplecount[t], datatype[t], filename[t]])
			hardwareeventtype[:] = []
	 		nor_hardwareeventsamplecount[:] = []
	 		nor_hardwareeventcount[:] = []
	 		datatype[:] = []
	 		filename[:] = []
		print name + "  Done" 
	print ("#")*150 #dir finished



def combine(folder1, folder2):
	directory1 = "C:\Users\\\Desktop\DUMMY_DATA2\DUMMY_DATA\\"+ folder1 +"_to_CSV_Normalized\\"
	directory2 = "C:\Users\\\Desktop\DUMMY_DATA2\DUMMY_DATA\\"+ folder2 +"_to_CSV_Normalized\\"
	directory3 = "C:\Users\\\Desktop\DUMMY_DATA2\DUMMY_DATA\Combined_Normalized\\"
	os.chdir(directory1)
	os.system("copy *.csv merged1.csv")
	os.chdir(directory2)
	os.system("copy *.csv merged2.csv")
	os.rename(directory1+ "merged1.csv", directory3+ "merged1.csv")
	os.rename(directory2+ "merged2.csv", directory3+ "merged2.csv")
	os.chdir(directory3)
	os.system("copy *.csv ALL_COMBINED_Normalized.csv")
	os.system("del /f merged1.csv")
	os.system("del /f merged2.csv")






#FINAL DONE 
def makedir(name):
	os.system("mkdir "+ name +"_column_to_CSV")
	os.system("mkdir "+ name +"_to_CSV_Normalized")
	os.system("mkdir Combined_Normalized")
	
def main():
	makedir(data1) # make folders
	makedir(data2)
	coltocsv(data1) #make col to sheets
	coltocsv(data2)
	normalizecolumns(data1)
	normalizecolumns(data2)
	combine(data1,data2)
	
main()






# testclean()
#another normalize function 
	# print ("*")*150
	# r = float(raw[-1] - raw[0])
	# # Normalize
	# normal = map(lambda x: (x - raw[0]) / r, raw)
	# print normal
