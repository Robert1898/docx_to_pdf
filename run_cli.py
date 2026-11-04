"""Punct de intrare pentru linia de comanda (docconv.exe)."""

import sys

from convertor.cli import main

if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
