import sys
from pathlib import Path

sys.path.append(str(Path(__file__).parent))

from factory import ConverterFactory

def main():
    if len(sys.argv) > 1:
        ConverterFactory.create_cli_converter().run()
    else:
        ConverterFactory.create_gui_converter().run()

if __name__ == "__main__":
    main()