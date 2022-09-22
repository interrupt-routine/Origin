import PyOrigin
import os
import xml.etree.ElementTree as ET
import re

def get_xml(text : str):
	print(text)
	lines = text.split('\r\n')
	idx = lines.index('[EXP_FILE]')
	print(lines)

	xml = ET.fromstring(lines[idx + 1])
	ET.indent(xml, space='\t')
	return xml

def extract_folder(folder : PyOrigin.CPyFolder) -> None:
	for pagebase in folder.PageBases():
		if pagebase.Type != PyOrigin.PGTYPE_WKS: # ignore non-worksheets
			continue

		short_name = pagebase.GetName()
		long_name = pagebase.GetLongName()
		# print('working on page: ' + long_name)
		# if not long_name.startswith('Ex'):
		# 	long_name = 'Em_' + long_name

		note = PyOrigin.Pages(short_name).Layers('Note')
		if note is None:
			print('this page does not have a Note sheet')
			continue

		text = note.Columns(0).GetData()[0]
# <SCD1 darkEnabled="1" blankEnabled="0" blankFile="" correctionEnabled="0"/>
# <SCD2 darkEnabled="1" blankEnabled="0" blankFile="" correctionEnabled="0"/></Correction>
		try:
			sdc1 = re.search('(SCD1) (darkEnabled="-?\d+") blankEnabled="0" blankFile="" (correctionEnabled="-?\d+")', text).groups()
			sdc2 = re.search('(SCD2) (darkEnabled="-?\d+") blankEnabled="0" blankFile="" (correctionEnabled="-?\d+")', text).groups()

			print('%s has %s %s %s     and    %s %s %s' % (long_name, sdc1[0], sdc1[1], sdc1[2], sdc2[0], sdc2[1], sdc2[2]))
		except AttributeError:
			print('%s is ill-formed' % long_name)





def main():
	folder = PyOrigin.ActiveFolder()

	print('\n\ncurrent folder: ' + folder.Path() + '\n\n')
	extract_folder(folder)

if __name__ == '__main__':
	main()