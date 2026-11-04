"""python -m convertor           -> interfata grafica
python -m convertor --cli ...   -> linia de comanda
"""

import sys

if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--cli":
        from convertor.cli import main as cli_main

        sys.exit(cli_main(sys.argv[2:]))
    from convertor.gui import main as gui_main

    sys.exit(gui_main(sys.argv[1:]))
