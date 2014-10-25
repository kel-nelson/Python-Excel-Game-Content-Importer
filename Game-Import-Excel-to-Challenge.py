import xlrd
import dicttoxml
import math

def process_rows(rows):
        return str(dicttoxml.dicttoxml(rows))


rows = []
workbook = xlrd.open_workbook('order.xlsx')
worksheet = workbook.sheet_by_name('GameInput-2014-1')
num_rows = worksheet.nrows - 1
num_cells = worksheet.ncols - 1
curr_row = -1
headers = ['title', 'answer1', 'answer2', 'answer3', 'answer4']
while curr_row < num_rows:
	curr_row += 1
	row = worksheet.row(curr_row)
	print('Row:', curr_row)
	curr_cell = -1
	last_row_found=False
	row = {}
	while curr_cell < num_cells:
		curr_cell += 1
		#Cell Types: 0=Empty, 1=Text, 2=Number, 3=Date, 4=Boolean, 5=Error, 6=Blank
		cell_type = worksheet.cell_type(curr_row, curr_cell)
		if(curr_cell == 1 and cell_type == 0):
			last_row_found=True
			break
		cell_value = worksheet.cell_value(curr_row, curr_cell)
		row[headers[curr_cell]] = cell_value
        
	if(last_row_found):
		break
	row["questionId"]=(curr_row+1)
	row["level"]=math.ceil((curr_row+1)/4)
	rows.append(row)

output_text = process_rows(rows)
f = open('data/game_output.xml', 'w')
f.write(output_text)
print("done")


