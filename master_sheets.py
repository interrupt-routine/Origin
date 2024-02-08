import sys
from enum import Enum
from typing import Dict, List
from datetime import datetime
import PyOrigin
# for type hints:
from PyOrigin import CPyOriginCollectionBase, CPyColumn, CPyWorksheet, CPyWorksheetPage, CPyPageBase, CPyFolder


X_START =  200
X_END   = 1000

X_NAME = 'Wavelength'

X_UNIT = 'nm'
Y_UNIT_NORMALIZED = 'a.u.'
Y_BASE_UNIT       = 'CPS'

NORMAL_LAYER_NAME = 'Data'
BATCH_LAYER_NAME  = 'Data_S1c'

PREFIX_BATCH = 'STACK'
PREFIX_NORM  = 'NORM'

class Mode(Enum):
# Default mode: extracts the second column (first Y column) from every
# worksheet in the current folder, normalizes them to [0; 1], sends
# them to a master sheet in the project's root folder.
# Both Emission and Excitation experiments are extracted (to separate masters) if present.
	AUTOMATIC   = 'automatic'
# Same as AUTOMATIC, but asks the user for a wavelength around which to normalize
# and for an experiment type (this is done by LabTalk and passed as paramater to Python).
# Only worksheets of the selected type are extracted.
	INTERACTIVE = 'interactive'
# Same as AUTOMATIC, but does not normalize, and uses a separate master sheet.
	TITRATION   = 'titration'
# Same as TITRATION, but extracts all Y columns, not just the first.
	BATCH       = 'batch'

class ExpType(Enum):
	EXCITATION = 'Ex'
	EMISSION   = 'Em'



def collections_count(self : CPyOriginCollectionBase) -> int:
	"""
	Patches a PyOrigin bug, GetCount() does not work with Folder.PageBases()
	"""
	count = 0
	for _ in self:
		count += 1
	return count

CPyOriginCollectionBase.GetCount = collections_count



def is_valid_page(sheet : CPyPageBase):
	long_name = sheet.GetLongName()
	return (sheet.Type is PyOrigin.PGTYPE_WKS) and (not long_name.startswith(PREFIX_NORM)) and (not long_name.startswith(PREFIX_BATCH))



def get_creation_date(page_short_name : str) -> datetime:
	VAR_NAME = 'creation_date'
	PyOrigin.LT_execute('string %s$=get_creation_date("%s")$;' % (VAR_NAME, page_short_name))
	datestring = PyOrigin.LT_get_str(VAR_NAME)
# format: "14/06/2023 07:44"
	parts = datestring.split(' ')
	(day, month, year) = [int(x) for x in parts[0].split('/')]
	(hours, minutes) = [int(x) for x in parts[1].split(':')]

	return datetime(year, month, day, hours, minutes)



class Column:
	def __init__(self, column_object : CPyColumn, x_start : int):
		"""
		Extracts the data from a PyOrigin Column object into a ColumnData instance
		"""
		self.long_name = column_object.GetLongName()
		self.comments =  column_object.GetComments()
		rows =           column_object.GetData()

		self.rows = [x for x in rows if x != '']
		offset = next(i for i, x in enumerate(rows) if x != '')
		self.x_start = x_start + offset
		self.x_end = x_start + len(self.rows) - 1

	def write_column(self, col_object : CPyColumn) -> None:
		col_object.SetComments(self.comments)
		col_object.SetLongName(self.long_name)
		col_object.SetUnits(Y_UNIT)
		col_object.SetType(PyOrigin.COLTYPE_DESIGN_Y)

		rows = self.rows
	# it is critical to insert empty strings in the empty cells
	# because of a bug in the GetData(start, end) function --> it returns a list
	# filled with None if the (start, end) range contains empty cells
	# at the beginning

		if self.x_start > X_START:
			rows = [''] * (self.x_start - X_START) + rows

		col_object.SetData(rows)

	def normalize(self):
		if len(self.rows) == 0:
			return

		min_val = min(self.rows)

		if MODE is Mode.INTERACTIVE:
			if self.x_start > NORM_WAVELENGTH or self.x_end < NORM_WAVELENGTH:
				raise IndexError(
					'column %s has range (%d, %d)nm but you selected a normalizing wavelength of %d' %
					(self.long_name, self.x_start, self.x_end, NORM_WAVELENGTH)
				)
			max_val = self.rows[NORM_WAVELENGTH - self.x_start]
		else:
			max_val = max(self.rows)

		self.rows = [(x - min_val) / (max_val - min_val) for x in self.rows]



class WorkSheet:
	def __init__(self, page : CPyWorksheetPage) -> None:
		self.name = page.GetName()
		self.long_name = page.GetLongName()
		self.creation_date = get_creation_date(self.name)

		worksheet = page.Layers(LAYER_NAME)
		if worksheet is None:
			print("error: page '%s' ('%s') does not have a %s layer" % (self.long_name, self.name, LAYER_NAME))
			return

		x_column = worksheet.Columns(0)
		x_values = x_column.GetData(0)
		self.x_start, self.end_x = int(x_values[0]), int(x_values[-1])

		self.y_columns = []

		if MODE is Mode.BATCH:
			for i in range(1, worksheet.GetColCount()):
				column = Column(worksheet.Columns(i), self.x_start)
				column.long_name = self.long_name + '-' + str(i)
				self.y_columns.append(column)
		else:
			column = Column(worksheet.Columns(1), self.x_start)
			if MODE is not Mode.TITRATION:
				column.normalize()
			column.long_name = self.long_name
			if MODE is Mode.INTERACTIVE:
				column.long_name += '__(%d)' % NORM_WAVELENGTH

			self.y_columns.append(column)

	def append_to_master_sheet(self, master_sheet : CPyWorksheet) -> None:
		"""
		Appends the columns into the master sheet.
		"""
		master_column_names = [col.GetLongName() for col in master_sheet.Columns()]
		columns = [column for column in self.y_columns if column.long_name not in master_column_names]

		if len(columns) == 0:
			return

		columns_count = master_sheet.GetColCount()
		if columns_count == 0:
		# creating the first (x) column
			master_sheet.InsertCol(0, X_NAME)
			x_column = master_sheet.Columns(0)
			x_column.SetUnits(X_UNIT)
			x_column.SetLongName(X_NAME)
			x_column.SetType(PyOrigin.COLTYPE_DESIGN_X)
			x_column.SetData(list(range(X_START, X_END + 1)))
			columns_count = 1

		# inserting the next (y) columns
		for i, column_data in enumerate(columns, columns_count):
			master_sheet.InsertCol(i, 'Y' + str(i))
			y_column = master_sheet.Columns(i)
			column_data.write_column(y_column)



def extract_folder(folder : CPyFolder) -> Dict[ExpType, List[WorkSheet]]:
	"""
	Extracts the individual worksheets from a folder.
	"""

	worksheets = {
		ExpType.EMISSION   : [],
		ExpType.EXCITATION : []
	}

	for pagebase in folder.PageBases():
		if not is_valid_page(pagebase):
			continue

		(page_name, page_longname) = (pagebase.GetName(), pagebase.GetLongName())

		exp_type = ExpType.EXCITATION if 'Ex' in page_name or 'Ex' in page_longname else ExpType.EMISSION
		if MODE is Mode.INTERACTIVE and exp_type is not EXP_TYPE:
			continue

		page = PyOrigin.Pages(page_name)
		worksheet = WorkSheet(page)

		print("page '%s' ( '%s' ) created %s has range (%d, %d) nm" %
			(page_name, page_longname, worksheet.creation_date, worksheet.x_start, worksheet.end_x)
		)

		worksheets[exp_type].append(worksheet)

	return worksheets



def create_worksheet(short_name : str, long_name : str) -> CPyWorksheet:
	page = PyOrigin.CreatePage(PyOrigin.PGTYPE_WKS, short_name, "", 1)
	page.SetLongName(long_name)

	worksheet = page.Layers(0)
	worksheet.SetName(NORMAL_LAYER_NAME)
	# they will not appear by default ...
	worksheet.SetLabelVisible(PyOrigin.LABEL_COMMENTS,  True)
	worksheet.SetLabelVisible(PyOrigin.LABEL_UNITS,     True)
	worksheet.SetLabelVisible(PyOrigin.LABEL_LONG_NAME, True)

	return worksheet



def make_master_sheet(exp_type : ExpType, prefix : str, data : Dict[ExpType, List[WorkSheet]]) -> None:
	worksheets = data[exp_type]
	if len(worksheets) == 0: # we do not create a master sheet if there is no data
		return

	# sorting the columns by long name (we want to do that *before* appending to the master sheets):
	worksheets.sort(key = lambda sheet : sheet.creation_date)

	start = PREFIX_BATCH if MODE is Mode.BATCH or MODE is Mode.TITRATION else PREFIX_NORM
	long_name = start + '_' + prefix
	if MODE is Mode.AUTOMATIC or MODE is Mode.INTERACTIVE:
		long_name += '_' + exp_type.value

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
		master_sheet = PyOrigin.Pages(short_name).Layers(NORMAL_LAYER_NAME)
	else:
		master_sheet = create_worksheet(exp_type.value + prefix, long_name)
		short_name = master_sheet.GetPage().GetName()
		print("created master sheet in the project's root with short name '%s' and long name '%s'"
			% (short_name, long_name)
		)

	cols_count = master_sheet.GetColCount()

	for worksheet in worksheets:
		worksheet.append_to_master_sheet(master_sheet)

	if master_sheet.GetColCount() == cols_count:
		print('All columns already existed in the master.')

def detect_batch_mode() -> bool:
	folder = PyOrigin.ActiveFolder()
	pagebases = (pagebase for pagebase in folder.PageBases() if is_valid_page(pagebase))
	pages = (PyOrigin.Pages(pagebase.GetName()) for pagebase in pagebases)

	try:
		page = next(pages)
		return page.Layers(BATCH_LAYER_NAME) is not None
	except StopIteration:
		return False



def main():
	folder = PyOrigin.ActiveFolder()
	folder_name = folder.GetName() # folders do not have long names

	parts = folder_name.split('_')
	prefix = parts[0] # TN76_DCM_... -> TN76
	if (MODE is Mode.TITRATION or MODE is Mode.BATCH) and len(parts) >= 2:
		prefix += '_' + parts[1] # TN76_DCM_... -> TN76_DCM


	print('=' * 80)
	print('current folder:\t' + folder.Path())
	print('=' * 80)
	print('\n')

	if folder.PageBases().GetCount() == 0:
		print('No worksheets found in the folder, nothing to do.')
		return

	worksheets = extract_folder(folder)

	if (len(worksheets[ExpType.EMISSION]) == 0) and (len(worksheets[ExpType.EXCITATION]) == 0):
		print('No suitable worksheets were found')
		return

	# moving to the project's root so that we can create sheets there
	PyOrigin.XF('pe_cd', {'path' : '/'})

	print('\n\n')

	if MODE is not Mode.INTERACTIVE or EXP_TYPE is ExpType.EMISSION:
		make_master_sheet(ExpType.EMISSION, prefix, worksheets)

	print('\n\n')

	if MODE is not Mode.INTERACTIVE or EXP_TYPE is ExpType.EXCITATION:
		make_master_sheet(ExpType.EXCITATION, prefix, worksheets)



if __name__ == '__main__':
	LAYER_NAME = NORMAL_LAYER_NAME
	Y_UNIT = Y_UNIT_NORMALIZED

	if len(sys.argv) == 3:
		MODE = Mode.INTERACTIVE
		NORM_WAVELENGTH = int(sys.argv[1])
		EXP_TYPE = ExpType.EMISSION if sys.argv[2] == 'Emission' else ExpType.EXCITATION
	elif len(sys.argv) == 2 and sys.argv[1] == 'titration':
		if detect_batch_mode():
			MODE = Mode.BATCH
			LAYER_NAME = BATCH_LAYER_NAME
		else:
			MODE = Mode.TITRATION
		Y_UNIT = Y_BASE_UNIT
	else:
		MODE = Mode.AUTOMATIC

	print("working in mode: " + str(MODE))
	main()