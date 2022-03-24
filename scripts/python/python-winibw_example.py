# -*- coding: utf-8 -*-

# External import
import python_init
import get_python_parameters # To retrieve Python-WinIBW parameters
from Abes_Apis_Interface.AbesXml import AbesXml # by Alexandre Faure (github.com/louxfaure)
# To connect to Abes's SudocXML webservice

# Stores Python-WinIBW parameters in pyth_par
pyth_par = get_python_parameters.main()

# Retrieves the PPN from js_to_pyth_file
with open(pyth_par["js_to_pyth_file"], "r", encoding="utf-8") as js_to_pyth:
    ppn = js_to_pyth.read()

# Connects to Abes's SudocXML webservice and retrieve the XML record
record = AbesXml(ppn)
if record.get_init_status() == "Succes":
    record = record.get_record # Returns the XML record as a string
else:
    record = "Could not return the record"

# Writes the result to pyth_to_js_file
with open(pyth_par["pyth_to_js_file"], "w", encoding="utf-8") as pyth_to_js:
    pyth_to_js.write(record)

# The python script ends, now we go back to the javascript script