"""
An experiment has 2 parts:

* EX1: Excitation 1 (Mono1)
* EM1: Emission 1 (Mono2)

They are identified by the attribute Device="Mono1" and Device = "Mono2", respectively.

The parameters for each these experiment are within the <Op> tags.
The parameters we want are the 4 <Op> elements
that have the attribute Command="2".

The first of these 4 elements is
In Emission:
    * the Park (for Excitation i.e. "Mono1")
    * the starting wavelength (for Emission i.e. "Mono2")
In Excitation it's swapped:
    * the starting wavelength (for Excitation i.e. "Mono1")
    * the Park  (for Emission i.e. "Mono2")


The following 3 elements are the
* Side Entrance Slit
* Side Exit Slit 
* First Intermediate Slit

in nanometers. They must be identical within each Experiment.
"""

#<EXPERIMENT>
#    <StartOps>

# Excitation:

#        <Op Command="2" Device="Mono1">
#            <Parameters>
#                <Param Type="2" Value=PARK />
# repeated 3 times:
#        <Op Command="2" Device="Mono1">
#            <Parameters>
#                <Param Type="2" Value=EXCITATION_SLIT />

# Emission:

#        <Op Command="2" Device="Mono2">
#            <Parameters>
#                <Param Type="2" Value=STARTING_WAVELENGTH />
# repeated 3 times:
#        <Op Command="2" Device="Mono2">
#            <Parameters>
#                <Param Type="2" Value=EMISSION_SLIT />

# Integration Time: (in Excitation there is also one with Device="SCD2")

#        <Op Command="5" Device="SCD1">
#            <Parameters>
#                <Param Type="3" Value="INTEGRATION_TIME" />

# Range:
# !!! --> In the experiment file embedded within the Notes,
# the <Param Type="2"> contains the ENDING wavelength instead of the STARTING

#    <ExpAxis>
#        <Axis Begin=STARTING_WAVELENGTH End=ENDING_WAVELENGTH>
#        <Operations>
# the Device is "Mono2" for Excitation and "Mono1" for Emission
#            <Op Command="2" Device="Mono2">
#                <Parameters>
#                    <Param Type="2" Value=STARTING_WAVELENGTH />



import xml.etree.ElementTree as ET
from collections import namedtuple
from math import sqrt
import enum
import os

EM_PARKS = (250, 275) + tuple(park for park in range(300, 550 + 10, 10))
EX_PARKS = tuple(park for park in range(350, 800 + 10, 10))
SLITS = ((1, 1), (2, 1), (3, 1), (3, 1.5), (3, 2), (4, 2), (5, 2), (6, 3), (6, 4), (8, 6), (10, 6), (10, 8), (15, 10))
INTEGRATION_TIMES = (0.1, 0.5, 1.0)
ROOT_DIR_NAME = 'Presets'

# the type is that of the <Param> elements, NOT of the <Op> elements
ElementCriteria = namedtuple('ElementCriteria', ['device', 'command', 'type_'])

excitation_criteria       = ElementCriteria(device = 'Mono1', type_ = '2', command = '2')
emission_criteria         = ElementCriteria(device = 'Mono2', type_ = '2', command = '2')
integration_time_criteria = ElementCriteria(device = None,    type_ = '3', command = '5')
end_wavelength_criteria   = ElementCriteria(device = None,    type_ = '2', command = '2')

class ExperimentType(enum.Enum):
    EXCITATION = 'Excitation'
    EMISSION   = 'Emission'

class AlwaysEqual:
    def __eq__(self, _):
        return True

dummy = AlwaysEqual() # this will act as a wildcard


def get_start_ops_params(ops : list[ET.Element], *, device : str, command : str, type_ : str) -> list[ET.Element]:
    type_   = dummy if type_   is None else type_
    device  = dummy if device  is None else device
    command = dummy if command is None else command

    return [
        next(elt for elt in params if elt.get('Type') == type_)
            for params in
        (op.find('Parameters').findall('Param') for op in ops
        if op.get('Device') == device and op.get('Command') == command)
    ]


class ExperimentXML:

    def __init__(self, filename : str, exp_type : ExperimentType):
        self.exp_type = exp_type
        self.tree = ET.parse(filename)
        ET.indent(self.tree, space = '') # makes the file a lot shorter
        self.axis = self.tree.find('ExpAxis').find('Axis')

        start_ops = self.tree.find('StartOps').findall('Op')
        axis_ops  = self.axis.find('Operations').findall('Op')

        self.excitation        = get_start_ops_params(start_ops, **excitation_criteria._asdict())
        self.emission          = get_start_ops_params(start_ops, **emission_criteria._asdict())
        self.integration_time  = get_start_ops_params(start_ops, **integration_time_criteria._asdict())
        self.start_wavelength  = get_start_ops_params(axis_ops,  **end_wavelength_criteria._asdict())[0]

    def generate_xml(self, *, ex_slit, em_slit, park, start_wavelength, end_wavelength, integration_time) -> str:
        args =\
        [ex_slit, em_slit, park, start_wavelength, end_wavelength, integration_time]
        [ex_slit, em_slit, park, start_wavelength, end_wavelength, integration_time] = [str(arg) for arg in args]

        match self.exp_type:
            case ExperimentType.EXCITATION:
                self.excitation[0].attrib['Value'] = start_wavelength
                self.emission  [0].attrib['Value'] = park
            case ExperimentType.EMISSION:
                self.excitation[0].attrib['Value'] = park
                self.emission  [0].attrib['Value'] = start_wavelength

        for elt in self.excitation[1:]:
            elt.attrib['Value'] = ex_slit

        for elt in self.emission[1:]:
            elt.attrib['Value'] = em_slit

        for elt in self.integration_time:
            elt.attrib['Value'] = integration_time

        self.axis.attrib['Begin'] = start_wavelength
        self.axis.attrib['End']   = end_wavelength

        self.start_wavelength.attrib['Value'] = start_wavelength

        xml_string = ET.tostring(self.tree.getroot(), encoding='unicode', xml_declaration=False, short_empty_elements=True)
        return xml_string.replace('\n', '')


def print_elems(elements : ET.Element | list[ET.Element]):
    def print_elem(elt : ET.Element) -> None:
        print(f"tag = {elt.tag},  attributes = {elt.attrib}, content = {elt.text}")

    if type(elements) is list:
        for elem in elements:
            print_elem(elem)
    else:
        print_elem(elements)



def select_range(exp_type : ExperimentType, park : int, ex_slit, em_slit) -> tuple[int, int]:
    def round_to_multiple(x, multiple):
        return round(float(x) / multiple) * multiple

    max_slit = max(ex_slit, em_slit)

    match exp_type:
        case ExperimentType.EXCITATION:
            if max_slit >= 10:
                S = 1.0
            elif max_slit >= 5:
                S = 0.9
            else:
                S = 0.7
            start = park / 2 + 20 * S * sqrt(em_slit + ex_slit)
            end   = park     - 20 * S * sqrt(em_slit + ex_slit)
        case ExperimentType.EMISSION:
            if max_slit >= 10:
                S = 1.0
            elif max_slit >= 5:
                S = 0.8
            else:
                S = 0.6
            start =     park + 20 * S * sqrt(em_slit + ex_slit)
            end   = 2 * park - 20 * S * sqrt(em_slit + ex_slit)

    return (round_to_multiple(start, 5), round_to_multiple(end, 5))



def generate_files(dir_path : str, exp_obj : ExperimentXML, *, em_slit, ex_slit, exp_type : ExperimentType, parks):
    for integration_time in INTEGRATION_TIMES:
        for park in parks:
            filename = f"{exp_type.value}_{park}_{ex_slit}_{em_slit}_{integration_time}.xml"
            path = f"{dir_path}/{integration_time}/{filename}"

            (start_wavelength, end_wavelength) = select_range(exp_type, park, ex_slit, em_slit)
            parameters = {
                'ex_slit' : ex_slit, 'em_slit' : em_slit,
                'park' : park, 'integration_time' : integration_time,
                'start_wavelength' : start_wavelength,
                'end_wavelength' : end_wavelength
            }
            xml_string = exp_obj.generate_xml(**parameters)

            with open(path, mode = 'w') as f:
                f.write(xml_string)

            print(f"{filename} has range {(start_wavelength, end_wavelength)}")

def mkdir(path):
    try:
        os.mkdir(path)
    except FileExistsError:
        pass


def main():
    EX_EXP = ExperimentXML('Excitation.xml', ExperimentType.EXCITATION)
    EM_EXP = ExperimentXML('Emission.xml',   ExperimentType.EMISSION)

    mkdir(ROOT_DIR_NAME)

    for slit in SLITS:
        em_slit, ex_slit = sorted(slit)

        dir_path = f"{ROOT_DIR_NAME}/{ex_slit}-{em_slit}"
        mkdir(dir_path)
        for time in INTEGRATION_TIMES:
            mkdir(f"{dir_path}/{time}")

        # Emission   : Ex >= Em
        exp_type = ExperimentType.EMISSION
        generate_files(dir_path, EM_EXP, em_slit=em_slit, ex_slit=ex_slit, exp_type=exp_type, parks=EM_PARKS)

        # Excitation : Em >= Ex
        exp_type = ExperimentType.EXCITATION
        em_slit, ex_slit = (ex_slit, em_slit)
        generate_files(dir_path, EX_EXP, em_slit=em_slit, ex_slit=ex_slit, exp_type=exp_type, parks=EX_PARKS)


if __name__ == '__main__':
	main()
