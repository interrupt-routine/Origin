import sys
from enum import Enum
from typing import Dict, List
import PyOrigin



X_START = 200
X_END   = 900
X_NAME = 'Wavelength'
X_UNIT = 'nm'
Y_UNIT = 'a.u.'


class Mode(Enum):
	INTERACTIVE = 'interactive'
	AUTOMATIC   = 'automatic'

class ExpType(Enum):
	EXCITATION = 'Ex'
	EMISSION   = 'Em'



def collections_count(collection : PyOrigin.CPyOriginCollectionBase) -> int:
	"""
	Patches a PyOrigin bug, GetCount() does not work with Folder.PageBases()
	"""
	count = 0
	for _ in collection:
		count += 1
	return count

PyOrigin.CPyOriginCollectionBase.GetCount = collections_count



class ColumnData:
	def __init__(self, column_object : PyOrigin.CPyColumn):
		"""
		Extracts the data from a PyOrigin Column object into a ColumnData instance
		"""
		self.long_name = column_object.GetLongName()
		self.comments =  column_object.GetComments()
		rows =           column_object.GetData()

		self.offset = next(i for i, x in enumerate(rows) if x != '')
		self.rows = [x for x in rows if x != '']

		self.start_x = self.end_x = 0



	def write_column(self, col_object : PyOrigin.CPyColumn) -> None:
		col_object.SetComments(self.comments)
		col_object.SetLongName(self.long_name)

		rows = self.rows
	# it is critical to insert empty strings in the empty cells
	# because of a bug in the GetData(start, end) function --> it returns a list
	# filled with None if the (start, end) range contains empty cells
	# at the beginning
		if self.start_x > X_START:
			rows = [''] * (self.start_x - X_START) + rows

		col_object.SetData(rows)



	def normalize(self):
		if len(self.rows) == 0:
			return

		min_val = min(self.rows)

		if MODE is Mode.INTERACTIVE:
			if self.start_x > NORM_WAVELENGTH or self.end_x < NORM_WAVELENGTH:
				raise IndexError(
					'column %s has range (%d, %d)nm but you selected a normalizing wavelength of %d' %
					(self.long_name, self.start_x, self.end_x, NORM_WAVELENGTH)
				)
			max_val = self.rows[NORM_WAVELENGTH - self.start_x]
		else:
			max_val = max(self.rows)

		self.rows = [(x - min_val) / (max_val - min_val) for x in self.rows]



def extract_worksheet(worksheet: PyOrigin.CPyWorksheet) -> ColumnData:
	"""
	Extracts the data and metadata from a worksheet.
	Only the second column (Y) is extracted, together with
	the corresponding range of X values (first column).
	"""
	column_data = ColumnData(worksheet.Columns(1))

# grabbing the x range for the y column
	x_col = worksheet.Columns(0)
	column_data.start_x = int(x_col.GetData(column_data.offset, column_data.offset)[0])
	column_data.end_x   = column_data.start_x + len(column_data.rows) - 1
	column_data.normalize()

	return column_data



def extract_master_sheet(master_sheet : PyOrigin.CPyWorksheet) -> List[ColumnData]:
	"""
	Extracts the data and metadata from a master sheet.
	All Y columns are extracted i.e. not including X, the first column.
	"""
	columns = []

	for i in range(1, master_sheet.GetColCount()):
		column_data = ColumnData(master_sheet.Columns(i))

		column_data.start_x = X_START + column_data.offset
		column_data.end_x = column_data.start_x + len(column_data.rows) - 1

		# print("column '%s' ( '%s' ) has range (%d, %d) nm" %
		# 	(master_sheet.Columns(i).GetName(), master_sheet.Columns(i).GetLongName(), column_data.start_x, column_data.end_x)
		# )

		columns.append(column_data)

	return columns



def write_to_master_sheet(master_sheet : PyOrigin.CPyWorksheet, columns : List[ColumnData]) -> None:
	"""
	Inserts the columns into the master sheet.
	"""
	for i in range(0, master_sheet.GetColCount()):
		master_sheet.DeleteCol(0)

# creating the first (x) column
	master_sheet.InsertCol(0, X_NAME)
	x_column = master_sheet.Columns(0)
	x_column.SetUnits(X_UNIT)
	x_column.SetLongName(X_NAME)
	x_column.SetType(PyOrigin.COLTYPE_DESIGN_X)
	x_column.SetData(list(range(X_START, X_END + 1)))

# sorting the columns by long name:
	columns.sort(key = lambda col : col.long_name)

# inserting the next (y) columns
	for i, column_data in enumerate(columns, 1):
		master_sheet.InsertCol(i, 'Y' + str(i))
		y_column = master_sheet.Columns(i)
		y_column.SetUnits(Y_UNIT)
		y_column.SetType(PyOrigin.COLTYPE_DESIGN_Y)
		column_data.write_column(y_column)



def extract_folder(folder : PyOrigin.CPyFolder) -> Dict[str, List[ColumnData]]:
	"""
	Extracts the individual worksheets from a folder. For each worksheet, the second column
	and the range of the first one are extracted.
	"""

	master_columns = {
		ExpType.EMISSION   : [],
		ExpType.EXCITATION : []
	}

	for pagebase in folder.PageBases():
		if pagebase.Type != PyOrigin.PGTYPE_WKS: # ignore non-worksheets
			continue

		page_name = pagebase.GetName()
		page_longname = pagebase.GetLongName()

		page = PyOrigin.Pages(page_name)

		worksheet = page.Layers('Data')
		if worksheet is None:
			continue # no Data layer

		exp_type = ExpType.EXCITATION if 'Ex' in page_name or 'Ex' in page_longname else ExpType.EMISSION

		if MODE is Mode.INTERACTIVE and exp_type is not EXP_TYPE:
			continue

		col = extract_worksheet(worksheet)
		print("page '%s' ( '%s' ) has range (%d, %d) nm" %
			(page_name, page_longname, col.start_x, col.end_x)
		)

		col.long_name = page_longname
		if MODE is Mode.INTERACTIVE:
			col.long_name += '__(%d)' % NORM_WAVELENGTH

		master_columns[exp_type].append(col)


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



def make_master_sheet(exp_type : ExpType, prefix : str, data : Dict[str, List[ColumnData]]) -> None:
	columns = data[exp_type]
	if len(columns) == 0: # we do not create a master sheet if there is no data
		return

	long_name = 'NORM' + '_' + prefix + '_' + exp_type.value

# - short names are silently truncated to 12 chars, special chars such as '-', '_' are silently removed
# - Pages() only works with short names, because they're unique per project, whereas long names are not
# --> we have to search for the long name, then get the short name

	page = next((
		page for page in PyOrigin.GetRootFolder().PageBases()
		if page.GetLongName() == long_name
	), None)

	if page is not None:
		short_name = page.GetName()
		print("found master sheet in the project's root with short name '%s' and long name '%s'" 
			% (short_name, long_name)
		)
		sheet = PyOrigin.Pages(short_name).Layers('Data')
		master_columns = extract_master_sheet(sheet)
		long_names = [column.long_name for column in master_columns]

		columns = [column for column in columns if column.long_name not in long_names]
		if len(columns) == 0:
			print('All columns already existed in the master sheet, nothing to do.')
			return

		columns += master_columns
	else:
		sheet = create_worksheet(exp_type.value + prefix, long_name)
		short_name = sheet.GetPage().GetName()
		print("created master sheet in the project's root with short name '%s' and long name '%s'"
			% (short_name, long_name)
		)

	write_to_master_sheet(sheet, columns)



def main():
	folder = PyOrigin.ActiveFolder()
	folder_name = folder.GetName() # folders do not have long names
	prefix = folder_name.split('_', 1)[0] # TN76_DCM_... -> TN76

	print('=' * 80)
	print('current folder:\t' + folder.Path())
	print('=' * 80)
	print('\n')

	if folder_name == PyOrigin.GetRootFolder().GetName():
		print("You called the script from the project's root, nothing to do.")
		return
	if folder.PageBases().GetCount() == 0:
		print('No worksheets found in the folder, nothing to do.')
		return

	data = extract_folder(folder)

	if (len(data[ExpType.EMISSION]) == 0) and (len(data[ExpType.EXCITATION]) == 0):
		print('No worksheets found in the folder, nothing to do.')
		return

	# moving to the project's root so that we can create sheets there
	PyOrigin.XF('pe_cd', {'path' : '/'})

	print('\n\n')

	if MODE is Mode.AUTOMATIC or EXP_TYPE is ExpType.EMISSION:
		make_master_sheet(ExpType.EMISSION, prefix, data)

	print('\n\n')

	if MODE is Mode.AUTOMATIC or EXP_TYPE is ExpType.EXCITATION:
		make_master_sheet(ExpType.EXCITATION, prefix, data)



if __name__ == '__main__':
	if (len(sys.argv)) == 3:
		MODE = Mode.INTERACTIVE
		NORM_WAVELENGTH = int(sys.argv[1])
		EXP_TYPE = ExpType.EMISSION if sys.argv[2] == 'Emission' else ExpType.EXCITATION
	else:
		MODE = Mode.AUTOMATIC

	main()