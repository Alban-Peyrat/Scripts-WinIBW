# -*- coding: utf-8 -*-
# Sets up WinIBW to use Python-WinIBW

WINIBW_ROOT_FOLDER = r"C:\oclcpica\WinIBW30"

import os
import shutil

files = ["clean_pyth_files.py",
        "get_python_parameters.py",
        "pythWinIBW.js",
        "python_parameters"]

# Sets up WinIBW's setup.js
with open(WINIBW_ROOT_FOLDER + r"\defaults\pref\setup.js", "r+", encoding="utf-8") as setupFile:
    content = setupFile.read()
    if not "python-winibw/pythWinIBW.js" in content:
        # Creates a copy of setup.js
        setupCopy = open(WINIBW_ROOT_FOLDER + r"\defaults\pref\setup_copy_by_python-winibw.js", "w", encoding="utf-8")
        setupCopy.write(content)
        setupCopy.close()
        # Modifies setup.js
        setupFile.seek(0, 0)
        setupFile.truncate()
        setupFile.write('pref("ibw.standardScripts.script.pythWinIBW", "resource:/SCOOP/scripts/python-winibw/pythWinIBW.js");\n')
        setupFile.write(content)

# Sets up python-winibw
os.makedirs(WINIBW_ROOT_FOLDER + r"\SCOOP\scripts\python-winibw", exist_ok=True)
fileDst = WINIBW_ROOT_FOLDER + r"\SCOOP\scripts\python-winibw\\"
for file in files:
    shutil.copyfile(file, fileDst+file)