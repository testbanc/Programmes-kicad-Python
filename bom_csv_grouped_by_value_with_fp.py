# -*- coding: utf-8 -*-
"""
    @package
    Output: CSV (comma-separated)
    Grouped By: Value, Footprint, DNP
    Sorted By: Ref
    Fields: Ref, Qnty, Value, Cmp name, Footprint, Description, Vendor, DNP

    Command line:
    python "D:\Programmes-kicad-Python\bom_csv_grouped_by_value_with_fp.py" "D:\Programmes-kicad\SYNOH052 carte etalon\SYNOH052 carte etalon.xml" "D:/Programmes-kicad/SYNOH052 carte etalon/SYNOH052 carte etalon.csv"
"""

import kicad_netlist_reader
import kicad_utils
import csv
import sys
from datetime import datetime

def fromNetlistText(aText):
    return aText if aText is not None else ''

net = kicad_netlist_reader.netlist(sys.argv[1])

try:
    with kicad_utils.open_file_writeUTF8(sys.argv[2], 'w') as f:
        out = csv.writer(f, lineterminator='\n', delimiter=';', quotechar='\"', quoting=csv.QUOTE_ALL)

        current_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        out.writerow(['Date:', current_date])
        out.writerow(['Programme python:', sys.argv[0]])
        out.writerow(['', ''])
        out.writerow(['', ''])  # Ligne vide pour séparer les en-têtes

        out.writerow(['Ref', 'Valeur', 'Fabriquant', 'Reference fabricant', 'Distributeur 1', 'Distributeur 2', 'Cmp name', 'Description', 'Quantite', 'Footprint', 'Documentation', 'DNP'])

        grouped = net.groupComponents()

        for group in grouped:
            for component in group:
                ref = fromNetlistText(component.getRef())
                value = fromNetlistText(component.getValue())
                cmp_name = fromNetlistText(component.getPartName())
                footprint = fromNetlistText(component.getFootprint())
                description = fromNetlistText(component.getDescription())
                fabriquant = fromNetlistText(component.getField("Fabriquant"))
                fabreference = fromNetlistText(component.getField("Fabreference"))
                documentation = fromNetlistText(component.getDatasheet())
                distributeur1 = fromNetlistText(component.getField("Distributeur 1"))
                distributeur2 = fromNetlistText(component.getField("Distributeur 2"))
                dnp = component.getDNPString()

                out.writerow([ref, value, fabriquant, fabreference, distributeur1, distributeur2, cmp_name, description, 1, footprint, documentation, dnp])

except IOError as e:
    print("Can't open output file for writing:", e, file=sys.stderr)
