"""Einstiegspunkt: startet den Einkauf Data Converter.

Start:            python data_converter_gui.py
Selbsttest:       python data_converter_gui.py --self-test
"""

import sys

from converter_app.gui import main

if __name__ == "__main__":
    main(sys.argv[1:])
