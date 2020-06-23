import xlsxwriter
import re
# import finditer

# Create file & worksheet
workbook = xlsxwriter.Workbook('result.xlsx')
worksheet = workbook.add_worksheet()

# Write some numbers, with row/column notation. 
# Note 0 starts.
f = open("Localizable.strings","r",encoding='UTF-8')

row = 0;
skip_lines = 7;

lines = f.readlines()
for line in lines:
	### Read Conditions, if not, skip ###
	#1 skip skip_lines
	if row < skip_lines: 
		row = row +1
		continue
	#2 skip \n
	if line == '\n':
		continue
	#3 skip contains '//'
	if line.find('//') != -1:
		continue
	
	### Deal with each line ###
	#1 remove ; at last
	line = line[:-2]
	#2 split
	results = re.split(r'=', line) #切等號
	#3 column0, column1
	count = 0
	column0 = '';
	column1 = '';
	for r in results:
		if count == 0:
			column0 = r.replace(' ','')
		if count == 1:
			r = r.replace(' "','')
			r = r.replace('"','')
			column1 = r
		count = count+1

	### write into file ###
	worksheet.write(row-skip_lines, 0, column0)
	worksheet.write(row-skip_lines, 1, column1)
	
	#next
	row = row +1

workbook.close()



# Functions

# v
# Create an new Excel file and add a worksheet.
# workbook = xlsxwriter.Workbook('demo.xlsx')
# worksheet = workbook.add_worksheet()

# v
# Write some simple text.
# worksheet.write('A1', 'Hello')


# v 
# 1. Add a bold format to use to highlight cells.
# 2. Text with formatting.
# bold = workbook.add_format({'bold': True})
# worksheet.write('A2', 'World', bold)

# ?
# Insert an image.
# worksheet.insert_image('B5', 'logo.png')

# ? 
# Widen the first column to make the text clearer.
# worksheet.set_column('A:A', 20)