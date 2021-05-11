#
# Example python script to generate a BOM from a KiCad generic netlist
#
# Example: Sorted and Grouped CSV BOM
#
"""
    @package
    Generate a csv BOM list.
    Components are sorted by ref and grouped by value
    Fields are (if exist)
    Qty, Reference(s), description, mfg1, mfg1pn, mfg2, mfg2pn
    
    To append additional lines to the generated BOM include a file
    "%O-aux.csv" in the project directory. The lines in the aux BOM
    will be appended to the generated BOM.

    Command line:
    python "pathToFile/bom_csv_grouped_by_value.py" "%I" "%O.csv"
"""

from __future__ import print_function

# Import the KiCad python helper module and the csv formatter
import kicad_netlist_reader
import csv
import sys
import os

desc = "description"
mfg1 = "mfg1"
mfg1pn = "mfg1pn"
mfg2 = "mfg2"
mfg2pn = "mfg2pn"
hmtFieldNames = [desc, mfg1, mfg1pn, mfg2, mfg2pn]

def toLower(str):
    return str.lower()

def getCaseInsensitiveField(comp, fieldName):
    """
    getCaseInsensitiveField gets a field from a component given a field name in a case insensitive way.
    """
    fieldNames = map(toLower, comp.getFieldNames())
    fieldName = fieldName.lower()
    if (fieldName not in fieldNames):
        return None
    return comp.getField(comp.getFieldNames()[fieldNames.index(fieldName)])

def equByHMTFields(comp, other):
    """
    equByHMTFields determines if two components are equal based on the HMT fields listed above.
    It returns None if the component's don't have values for any of hmtFieldNames
    """
    compFieldNames = set(map(toLower, comp.getFieldNames())) & set(hmtFieldNames)
    otherFieldNames = set(map(toLower, other.getFieldNames())) & set(hmtFieldNames)
    allFieldNames = compFieldNames | otherFieldNames
    if len(allFieldNames) == 0:
        return None
    result = True
    for name in allFieldNames:
        compField = getCaseInsensitiveField(comp, name)
        otherField = getCaseInsensitiveField(other, name)
        result = result and compField == otherField
    return result

def myEqu(self, other):
    """myEqu is a more advanced equivalence function for components which is
    used by component grouping. Normal operation is to group components based
    on their value and footprint.

    In this example of a custom equivalency operator we compare the, description field (if existing),
    then value, the part name and the footprint.
    """
    result = equByHMTFields(self, other)
    if result != None:
        return result
    elif self.getValue() != other.getValue():
        result = False
    elif self.getPartName() != other.getPartName():
        result = False
    elif self.getFootprint() != other.getFootprint():
        result = False

    return result

# Override the component equivalence operator - it is important to do this
# before loading the netlist, otherwise all components will have the original
# equivalency operator.
kicad_netlist_reader.comp.__eq__ = myEqu

if len(sys.argv) != 3:
    print("Usage ", __file__, "<generic_netlist.xml> <output.csv>", file=sys.stderr)
    sys.exit(1)


# Generate an instance of a generic netlist, and load the netlist tree from
# the command line option. If the file doesn't exist, execution will stop
net = kicad_netlist_reader.netlist(sys.argv[1])

# Open a file to write to, if the file cannot be opened output to stdout
# instead
try:
    f = open(sys.argv[2], 'w')
except IOError:
    e = "Can't open output file for writing: " + sys.argv[2]
    print(__file__, ":", e, file=sys.stderr)
    f = sys.stdout

# subset the components to those wanted in the BOM, controlled
# by <configure> block in kicad_netlist_reader.py
components = net.getInterestingComponents()

filteredComponents = []
for component in components:
    exclude = getCaseInsensitiveField(component, 'exclude')
    if exclude is None:
        filteredComponents.append(component)

components = filteredComponents

columns = ['Qty', 'Reference(s)', 'description', 'mfg1', 'mfg1pn', 'mfg2', 'mfg2pn']

# Create a new csv writer object to use as the output formatter
out = csv.writer(f, lineterminator='\n', delimiter=',', quotechar='\"', quoting=csv.QUOTE_ALL)

# override csv.writer's writerow() to support encoding conversion (initial encoding is utf8):
def writerow(acsvwriter, columns):
    utf8row = []
    for col in columns:
        utf8row.append(str(col))  # currently, no change
    acsvwriter.writerow(utf8row)

# Output all the interesting components individually first:
row = []

# header
writerow(out, columns)                   # reuse same columns

# Get all of the components in groups of matching parts + values
# (see kicad_netlist_reader.py)
grouped = net.groupComponents(components)

# Output component information organized by group, aka as collated:
item = 0
for group in grouped:
    del row[:]
    refs = ""

    # Add the reference of every component in the group and keep a reference
    # to the component so that the other data can be filled in once per group
    for component in group:
        if len(refs) > 0:
            refs += ", "
        refs += component.getRef()
        c = component

    row.append(len(group))
    row.append(refs)

    # from column 2 upwards, use the fieldnames to grab the data
    for field in columns[2:]:
        row.append(net.getGroupField(group, field))

    writerow(out, row)

#append rows from auxiliary bom
auxBomFileName = os.path.splitext(sys.argv[2])[0] + "-aux.csv"
try:
    with(open(auxBomFileName, 'r')) as auxBomFile:
        print("aux bom found at " + auxBomFileName)
        csvRows = csv.DictReader(auxBomFile)
        fieldnames = csvRows.fieldnames
        if (not set(columns).issubset(fieldnames)):
            print(__file__, ": ", "auxiliary bom must contain ", columns)
        else:
            for row in csvRows:
                def rowValueForName(name):
                    return row[name]
                writerow(out, map(rowValueForName, columns))
    
except IOError:
    e = "No auxiliary bom found at " + auxBomFileName
    print(__file__, ":", e, file=sys.stdout)


f.close()
