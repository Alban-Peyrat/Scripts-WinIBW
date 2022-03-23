# -*- coding: utf-8 -*-

import os
# Internal import
import get_python_parameters

# Removes pyth_to_js_file and js_to_pyth_file

param = get_python_parameters.main()

try:
    os.remove(param["pyth_to_js_file"])
except:
    pass

try:
    os.remove(param["js_to_pyth_file"])
except:
    pass
