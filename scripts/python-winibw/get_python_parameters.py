# -*- coding: utf-8 -*-

import json

def main():
    """Returns Python-WinIBW parameters as a dictionnary."""
    with open(r"C:\oclcpica\WinIBW30\SCOOP\scripts\python-winibw\python_parameters", encoding="utf-8") as fJson:
        pythonParam = json.load(fJson)
    return pythonParam