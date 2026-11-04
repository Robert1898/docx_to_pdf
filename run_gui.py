"""Punct de intrare pentru interfata grafica (si pentru exe-ul PyInstaller)."""

import sys

from convertor.gui import main

if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
