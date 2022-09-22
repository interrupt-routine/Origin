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
#                <Param Type="2" Value= PARK />
# repeated 3 times:
#        <Op Command="2" Device="Mono1">
#            <Parameters>
#                <Param Type="2" Value=EXCITATION_SLIT />

# Emission:

#        <Op Command="2" Device="Mono2">
#            <Parameters>
#                <Param Type="2" Value= STARTING_WAVELENGTH />
# repeated 3 times:
#        <Op Command="2" Device="Mono2">
#            <Parameters>
#                <Param Type="2" Value=EMISSION_SLIT />

# Integration Time: (in Excitation there is also one with Device="SCD2")

#        <Op Command="5" Device="SCD1">
#            <Parameters>
#                <Param Type="3" Value="0.1" />

# Range:

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

# the type is that of the <Param> elements, NOT of the <Op> elements
ElementCriteria = namedtuple('ElementCriteria', ['device', 'command', 'type_'])

excitation_criteria       = ElementCriteria(device = 'Mono1', type_ = '2', command = '2')
emission_criteria         = ElementCriteria(device = 'Mono2', type_ = '2', command = '2')
integration_time_criteria = ElementCriteria(device = None,    type_ = '3', command = '5')
excitation_criteria       = ElementCriteria(device = 'Mono1', type_ = '2', command = '2')
end_wavelength_criteria   = ElementCriteria(device = None,    type_ = '2', command = '2')

class AlwaysEqual:
    def __eq__(self, _):
        return True

dummy = AlwaysEqual() # this will act as a wildcard

def get_start_ops_params(ops : list[ET.Element], *, device : str, command : str, type_ : str) -> list[ET.Element]:
    type_   = dummy if type_   is None else type_
    device  = dummy if device  is None else device
    command = dummy if command is None else command

    return  [
        next(elt for elt in params if elt.get('Type') == type_)
            for params in
        (op.find('Parameters').findall('Param') for op in ops
        if op.get('Device') == device and op.get('Command') == command)
    ]

def print_elem(elt : ET.Element) -> None:
    print(f"tag = {elt.tag},  attributes = {elt.attrib}, content = {elt.text}")

def print_elems(elements):
    for elem in elements:
        print_elem(elem)


tree = ET.parse('Emission.xml')

axis = tree.find('ExpAxis').find('Axis')

START_OPS = tree.find('StartOps').findall('Op')
AXIS_OPS  = axis.find('Operations').findall('Op')



excitation_elts        = get_start_ops_params(START_OPS, **excitation_criteria._asdict())
emission_elts          = get_start_ops_params(START_OPS, **emission_criteria._asdict())
integration_time_elts  = get_start_ops_params(START_OPS, **integration_time_criteria._asdict())
end_wvlgt_elts         = get_start_ops_params(AXIS_OPS,  **end_wavelength_criteria._asdict())

# print_elems(excitation_elts)
# print_elems(emission_elts)
# print_elems(integration_time_elts)
# print_elems(end_wvlgt_elts)
# print_elem(axis)




PARK_START = 250
PARK_END   = 450

SLIT_START = 10
SLIT_END   = 35

def round_to_multiple(x, multiple):
    return round(float(x) / multiple) * multiple

folders, count  = 0, 0
for ex_slit in range(SLIT_START, SLIT_END + 1, 5):
    for em_slit in range(SLIT_START, ex_slit + 1, 5):
        folders += 1
        for park in range(PARK_START, PARK_END + 1, 10):
            range_start = 1.0 * park + 20.0 * 0.6 * sqrt((em_slit + ex_slit) / 10.0)
            range_end   = 2.0 * park - 20.0 * 0.6 * sqrt((em_slit + ex_slit) / 10.0)
            range_start, range_end = [round_to_multiple(range_start, 5), round_to_multiple(range_end, 5)]

            count += 1
            print(f'ex = {ex_slit / 10}, em = {em_slit / 10}, park = {park}, range = ({range_start}, {range_end})')

print(folders)