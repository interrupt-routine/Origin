import PyOrigin
import re
from typing import Dict
from collections import namedtuple

Parameters = namedtuple('Parameters', [
	'experiment_type',
	'park',
	'integration_time',
	'excitation_slit',
	'emission_slit',
])

FLOAT_PATTERN = '(\\d+(?:\\.\\d+)?)'
SLIT_PATTERN = '.*\n(?:.+\n)*Side Entrance Slit: %s nmBandpass' % FLOAT_PATTERN

def get_parameter(pattern : str, key : str, string: str, parameters : Dict):
	match = re.search(pattern = pattern, string = string)
	if match:
		parameters[key] = match.groups()[0]
	else:
		raise KeyError('could not find the following experimental parameter: ' + key)


def parse_experiment(text : str) -> Parameters:
	text = text.replace('\r\n', '\n')

	PATTERNS = [
		('experiment_type', 'Experiment Type:.*(Emission|Excitation)].*'),
		('park', 'Park: (\\d+)'),
		('integration_time', 'Integration Time: %ss' % FLOAT_PATTERN),
		('excitation_slit', 'Excitation' + SLIT_PATTERN),
		('emission_slit', 'Emission' + SLIT_PATTERN),
	]

	parameters = {}

	for key, pattern in PATTERNS:
		get_parameter(pattern, key, text, parameters)

	return Parameters(**parameters)

def rename_pages(folder : PyOrigin.CPyFolder) -> None:
	for pagebase in folder.PageBases():
		if pagebase.Type != PyOrigin.PGTYPE_WKS: # ignore non-worksheets
			continue

		short_name = pagebase.GetName()
		print('working on page: ' + short_name)

		note = PyOrigin.Pages(short_name).Layers('Note')
		if note is None:
			print('this page does not have a Note sheet')
			continue

		text = note.Columns(0).GetData()[0]

		try:
			parameters = parse_experiment(text)
		except KeyError as error:
			print(error)

		print(parameters)


def main():
	folder = PyOrigin.ActiveFolder()
	print('\n\ncurrent folder: ' + folder.Path() + '\n\n')

	rename_pages(folder)


if __name__ == '__main__':
	main()