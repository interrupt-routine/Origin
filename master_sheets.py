from typing import Dict, List # redundant in 3.9+, dict and list can be used directly, but we target older versions
import PyOrigin



X_START = 200
X_END   = 900
X_NAME = 'Wavelength'
X_UNIT = 'nm'
Y_UNIT = 'a.u.'


class ColumnData:
	start_x = 0
	end_x = 0

	def __init__(self, name, long_name, comments, units, data):
		self.name = name
		self.long_name = long_name
		self.comments = comments
		self.units = units
		self.data = data



def collections_count(collection : PyOrigin.CPyOriginCollectionBase):
	""" Necessary to bypass a PyOrigin bug, GetCount() does not work. """
	count = 0
	for _ in collection:
		count += 1
	return count



def normalize(array : List) -> List:
	if len(array) == 0:
		return []
	min_val = min(array)
	max_val = max(array)
	return [(x - min_val) / (max_val - min_val) for x in array]



def extract_column(column_object : PyOrigin.CPyColumn) -> ColumnData:
	"""
	Extracts the data from a PyOrigin Column object into a ColumnData instance
	"""
	name =      column_object.GetName()
	long_name = column_object.GetLongName()
	comments =  column_object.GetComments()
	data =      column_object.GetData()
	units =     Y_UNIT

	data = normalize([x for x in data if x != ''])

	return ColumnData(name, long_name, comments, units, data)



def extract_worksheet(worksheet: PyOrigin.CPyWorksheet) -> ColumnData:
	"""
	Extracts the data and metadata from a worksheet.
	Only the second column (Y) is extracted, together with
	the corresponding range of X values (first column).
	"""
	column_data = extract_column(worksheet.Columns(1))
# grabbing the x range for the y column
	x_data = worksheet.Columns(0).GetData()
	column_data.start_x = int(min(x_data))
	column_data.end_x   = int(max(x_data))

	return column_data



def write_column(col_object : PyOrigin.CPyColumn, col_data : ColumnData) -> None:
	col_object.SetComments(col_data.comments)
	col_object.SetUnits(   col_data.units)
	col_object.SetLongName(col_data.long_name)

	data = col_data.data
	if col_data.start_x > X_START:
		data = [''] * (col_data.start_x - X_START) + data
	col_object.SetData(data)



def write_columns(worksheet : PyOrigin.CPyWorksheet, columns : List[ColumnData]) -> None:
	cols_count = worksheet.GetColCount()

# creating the first (x) column, if it does not already exists
	if cols_count == 0:
		worksheet.InsertCol(0, X_NAME)
		first_column = worksheet.Columns(0)
		first_column.SetUnits(X_UNIT)
		first_column.SetLongName(X_NAME)
		first_column.SetType(PyOrigin.COLTYPE_DESIGN_X)
		first_column.SetData(list(range(X_START, X_END + 1)))
		cols_count += 1

# inserting the next (y) columns
	for column_data in columns:
		worksheet.InsertCol(cols_count, 'Y' + str(cols_count))
		column = worksheet.Columns(cols_count)
		column.SetType(PyOrigin.COLTYPE_DESIGN_Y)
		write_column(column, column_data)
		cols_count += 1



def extract_folder(folder : PyOrigin.CPyFolder) -> Dict[str, List[ColumnData]]:
	"""
	Extracts the individual worksheets from a folder. For each worksheet, the second column
	and the range of the first one are extracted.
	"""

	master_columns = {'Ex' : [], 'Em' : []}

	for pagebase in folder.PageBases():
		if pagebase.Type != PyOrigin.PGTYPE_WKS: # ignore non-worksheets
			continue

		page_name = pagebase.GetName()
		page_longname = pagebase.GetLongName()

		page = PyOrigin.Pages(page_name)

		worksheet = page.Layers('Data')

		col = extract_worksheet(worksheet)
		print("page '%s' ( '%s' ) has range (%d, %d) nm" %
			(page_name, page_longname, col.start_x, col.end_x)
		)

		col.long_name = page_longname
		col.name = page_name

		if 'Ex' in page_name or 'Ex' in page_longname:
			master_columns['Ex'].append(col)
		else:
			master_columns['Em'].append(col)

	return master_columns



def create_worksheet(short_name : str, long_name : str) -> PyOrigin.CPyWorksheet:
	page = PyOrigin.CreatePage(PyOrigin.PGTYPE_WKS, short_name, "", 1)
	page.SetLongName(long_name)

	worksheet = page.Layers(0)
	worksheet.SetName('Data')
	# they will not appear by default ...
	worksheet.SetLabelVisible(PyOrigin.LABEL_COMMENTS,  True)
	worksheet.SetLabelVisible(PyOrigin.LABEL_UNITS,     True)
	worksheet.SetLabelVisible(PyOrigin.LABEL_LONG_NAME, True)

	return worksheet



def make_master_sheet(id : str, prefix : str, data : Dict[str, List[ColumnData]]) -> None:
	columns = data[id]

	long_name = 'NORM' + '_' + prefix + '_' + id

# - short names are silently truncated to 12 chars, special chars such as '-', '_' are silently removed
# - Pages() only works with short names, because they're unique per project, whereas long names are not
# --> we have to search for the long name, then get the short name

	for page in PyOrigin.GetRootFolder().PageBases():
		if page.GetLongName() == long_name:
			short_name = page.GetName()
			print("found master sheet in the project's root with short name '%s' and long name '%s'" % (short_name, long_name))
			sheet = PyOrigin.Pages(short_name).Layers('Data')
			break
	else:
		sheet = create_worksheet(id + prefix, long_name)
		short_name = sheet.GetPage().GetName()
		print("created master sheet in the project's root with short name '%s' and long name '%s'" % (short_name, long_name))

	write_columns(sheet, columns)



def main():
	folder = PyOrigin.ActiveFolder()
	folder_name = folder.GetName() # folders do not have long names
	prefix = folder_name.split('_', 1)[0] # TN76_DCM_... -> TN76

	print('=' * 80)
	print('\ncurrent folder:\t' + folder.Path())
	print('=' * 80)

	if folder_name == PyOrigin.GetRootFolder().GetName():
		print("You called the script from the project's root, nothing to do.")
		return
	if collections_count(folder.PageBases()) == 0:
		print('No worksheets found in the folder, nothing to do.')
		return

	data = extract_folder(folder)

	# moving to the project's root so that we can create sheets there
	PyOrigin.XF('pe_cd', {'path' : '/'})

	print('\n\n')

	make_master_sheet('Em', prefix, data)

	print('\n\n')

	make_master_sheet('Ex', prefix, data)



if __name__ == '__main__':
	main()